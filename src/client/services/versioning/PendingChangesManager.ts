/**
 * Pending Changes Manager
 * 
 * Manages pending changes that require user approval before being committed.
 * Works with the existing version history system to provide an approval workflow.
 */

import { v4 as uuidv4 } from 'uuid';
import { 
  PendingChange, 
  PendingChangeStatus, 
  CreatePendingChangeOptions,
  PendingChangeResult
} from '../../models/PendingChangesModels';
import { VersionHistoryService } from './VersionHistoryService';
import { ActionRecorder } from './ActionRecorder';
import { ExcelOperation, ExcelOperationType } from '../../models/ExcelOperationModels';
import { UndoHandlers } from './UndoHandlers';
import { AffectedRange, VersionEventType } from '../../models/VersionModels';

/**
 * Service for managing pending changes that require user approval
 */
export class PendingChangesManager {
  private pendingChanges: Map<string, PendingChange> = new Map();
  private workbookPendingChanges: Map<string, string[]> = new Map(); // workbookId -> changeIds
  private versionHistoryService: VersionHistoryService;
  private actionRecorder: ActionRecorder;
  private undoHandlers: UndoHandlers;
  
  // Storage keys
  private storageKeyPrefix = 'excel-addin-pending-';
  private pendingChangesKey = 'changes';
  private workbookPendingChangesKey = 'workbook-changes';
  
  constructor(versionHistoryService: VersionHistoryService, actionRecorder: ActionRecorder) {
    this.versionHistoryService = versionHistoryService;
    this.actionRecorder = actionRecorder;
    this.undoHandlers = new UndoHandlers();
    this.loadFromStorage();
  }
  
  /**
   * Load pending changes data from localStorage
   */
  private loadFromStorage(): void {
    try {
      const startTime = performance.now();
      console.log(`🔍 [PendingChangesManager] Loading pending changes from localStorage...`);
      
      // Load pending changes
      const changesJson = localStorage.getItem(`${this.storageKeyPrefix}${this.pendingChangesKey}`);
      if (changesJson) {
        const changesArray = JSON.parse(changesJson) as PendingChange[];
        changesArray.forEach(change => {
          this.pendingChanges.set(change.id, change);
        });
        console.log(`📚 [PendingChangesManager] Loaded ${changesArray.length} pending changes from localStorage`);
      }
      
      // Load workbook pending changes mapping
      const workbookChangesJson = localStorage.getItem(`${this.storageKeyPrefix}${this.workbookPendingChangesKey}`);
      if (workbookChangesJson) {
        const workbookChangesMap = JSON.parse(workbookChangesJson) as Record<string, string[]>;
        Object.entries(workbookChangesMap).forEach(([workbookId, changeIds]) => {
          this.workbookPendingChanges.set(workbookId, changeIds);
        });
        console.log(`📁 [PendingChangesManager] Loaded workbook pending changes mapping for ${this.workbookPendingChanges.size} workbooks`);
      }
      
      const endTime = performance.now();
      console.log(`✅ [PendingChangesManager] Loaded pending changes in ${(endTime - startTime).toFixed(2)}ms`);
    } catch (error) {
      console.error(`❌ [PendingChangesManager] Error loading pending changes from localStorage:`, error);
    }
  }
  
  /**
   * Save pending changes data to localStorage
   */
  private saveToStorage(): void {
    try {
      const startTime = performance.now();
      console.log(`💾 [PendingChangesManager] Saving data to localStorage...`);
      
      // Save pending changes
      const changesArray = Array.from(this.pendingChanges.values());
      localStorage.setItem(`${this.storageKeyPrefix}${this.pendingChangesKey}`, JSON.stringify(changesArray));
      console.log(`📁 [PendingChangesManager] Saving ${changesArray.length} pending changes to localStorage`);
      
      // Save workbook pending changes mapping
      const workbookChangesMap: Record<string, string[]> = {};
      this.workbookPendingChanges.forEach((changeIds, workbookId) => {
        workbookChangesMap[workbookId] = changeIds;
      });
      localStorage.setItem(`${this.storageKeyPrefix}${this.workbookPendingChangesKey}`, JSON.stringify(workbookChangesMap));
      console.log(`📂 [PendingChangesManager] Saving workbook pending changes mapping for ${this.workbookPendingChanges.size} workbooks`);
      
      const endTime = performance.now();
      console.log(`✅ [PendingChangesManager] Successfully saved all data to localStorage in ${(endTime - startTime).toFixed(2)}ms`);
    } catch (error) {
      console.error(`❌ [PendingChangesManager] Error saving pending changes to localStorage:`, error);
    }
  }
  
  /**
   * Create a new pending change
   * @param options Options for creating the pending change
   * @returns The created pending change
   */
  public createPendingChange(options: CreatePendingChangeOptions): PendingChange {
    const { workbookId, operation, beforeState, commandId, description } = options;
    
    // Generate a unique ID for the pending change
    const id = uuidv4();
    
    // Create the pending change object
    const pendingChange: PendingChange = {
      id,
      workbookId,
      operation,
      beforeState,
      status: PendingChangeStatus.PENDING,
      timestamp: Date.now(),
      affectedRanges: this.extractAffectedRangesFromOperation(operation),
      commandId,
      description: description || this.generateChangeDescription(operation)
    };
    
    // Store the pending change
    this.pendingChanges.set(id, pendingChange);
    
    // Add to workbook pending changes mapping
    if (!this.workbookPendingChanges.has(workbookId)) {
      this.workbookPendingChanges.set(workbookId, []);
    }
    this.workbookPendingChanges.get(workbookId)!.push(id);
    
    // Save to storage
    this.saveToStorage();
    
    console.log(`✅ [PendingChangesManager] Created pending change: ${id} for workbook: ${workbookId}`);
    return pendingChange;
  }
  
  /**
   * Generate a description for a pending change based on the operation
   * @param operation The Excel operation
   * @returns A human-readable description of the change
   */
  private generateChangeDescription(operation: ExcelOperation): string {
    const op = operation as any;
    // Get the operation type as a string
    const opType = (op.op || 'unknown');
    
    switch (opType) {
      case ExcelOperationType.SET_VALUE:
        return `Set value in ${op.target}`;
      case ExcelOperationType.ADD_FORMULA:
        return `Add formula to ${op.target}`;
      case ExcelOperationType.FORMAT_RANGE:
        return `Format range ${op.range}`;
      case ExcelOperationType.CREATE_SHEET:
        return `Create sheet "${op.name}"`;
      case ExcelOperationType.DELETE_SHEET:
        return `Delete sheet "${op.name}"`;
      case ExcelOperationType.RENAME_SHEET:
        return `Rename sheet from "${op.sheet}" to "${op.name}"`;
      case ExcelOperationType.SET_ACTIVE_SHEET:
        return `Set active sheet to "${op.name}"`;
      case ExcelOperationType.CLEAR_RANGE:
        return `Clear range ${op.range}`;
      case ExcelOperationType.CREATE_TABLE:
        return `Create table in range ${op.range}`;
      case ExcelOperationType.SORT_RANGE:
        return `Sort range ${op.range}`;
      case ExcelOperationType.FILTER_RANGE:
        return `Filter range ${op.range}`;
      default:
        return `${opType} operation`;
    }
  }
  
  /**
   * Get all pending changes for a workbook
   * @param workbookId The workbook ID
   * @returns Array of pending changes for the workbook
   */
  public getPendingChangesForWorkbook(workbookId: string): PendingChange[] {
    const changeIds = this.workbookPendingChanges.get(workbookId) || [];
    return changeIds
      .map(id => this.pendingChanges.get(id))
      .filter(change => change !== undefined && change.status === PendingChangeStatus.PENDING) as PendingChange[];
  }
  
  /**
   * Get a pending change by ID
   * @param changeId The pending change ID
   * @returns The pending change or undefined if not found
   */
  public getPendingChange(changeId: string): PendingChange | undefined {
    return this.pendingChanges.get(changeId);
  }
  
  /**
   * Accept a pending change, committing it to the version history
   * @param changeId The pending change ID
   * @returns Result of the accept operation
   */
  public acceptPendingChange(changeId: string): PendingChangeResult {
    const change = this.pendingChanges.get(changeId);
    
    if (!change) {
      return {
        success: false,
        message: `Pending change with ID ${changeId} not found`,
        changeId
      };
    }
    
    if (change.status !== PendingChangeStatus.PENDING) {
      return {
        success: false,
        message: `Pending change with ID ${changeId} is already ${change.status}`,
        changeId
      };
    }
    
    try {
      // Update the status to accepted
      change.status = PendingChangeStatus.ACCEPTED;
      this.pendingChanges.set(changeId, change);
      
      // Record the operation in the version history
      this.versionHistoryService.recordAction(
        change.workbookId,
        change.operation,
        change.beforeState,
        undefined // No metadata
      );
      
      // Save to storage
      this.saveToStorage();
      
      console.log(`✅ [PendingChangesManager] Accepted pending change: ${changeId}`);
      return {
        success: true,
        message: `Successfully accepted change: ${change.description}`,
        changeId
      };
    } catch (error) {
      console.error(`❌ [PendingChangesManager] Error accepting pending change ${changeId}:`, error);
      return {
        success: false,
        message: `Error accepting change: ${error instanceof Error ? error.message : String(error)}`,
        changeId
      };
    }
  }
  
  /**
   * Reject a pending change, reverting the changes made
   * @param changeId The pending change ID
   * @returns Result of the reject operation
   */
  public async rejectPendingChange(changeId: string): Promise<PendingChangeResult> {
    const change = this.pendingChanges.get(changeId);
    
    if (!change) {
      return {
        success: false,
        message: `Pending change with ID ${changeId} not found`,
        changeId
      };
    }
    
    if (change.status !== PendingChangeStatus.PENDING) {
      return {
        success: false,
        message: `Pending change with ID ${changeId} is already ${change.status}`,
        changeId
      };
    }
    
    try {
      // Update the status to rejected
      change.status = PendingChangeStatus.REJECTED;
      this.pendingChanges.set(changeId, change);
      
      // Create a workbook action from the pending change
      const workbookAction = {
        id: changeId,
        workbookId: change.workbookId,
        timestamp: change.timestamp,
        type: VersionEventType.RangeOperation,
        operation: change.operation,
        description: change.description,
        affectedRanges: change.affectedRanges.map(rangeStr => {
          const { sheet, address } = this.parseReference(rangeStr);
          return {
            sheetName: sheet,
            range: address,
            type: 'range' as 'cell' | 'range' | 'sheet' | 'table' | 'chart'
          };
        }),
        beforeState: change.beforeState
      };
      
      // Revert the changes using the undo handler
      await Excel.run(async (context) => {
        await this.undoHandlers.undoAction(context, workbookAction);
      });
      
      // Save to storage
      this.saveToStorage();
      
      console.log(`✅ [PendingChangesManager] Rejected pending change: ${changeId}`);
      return {
        success: true,
        message: `Successfully rejected change: ${change.description}`,
        changeId
      };
    } catch (error) {
      console.error(`❌ [PendingChangesManager] Error rejecting pending change ${changeId}:`, error);
      return {
        success: false,
        message: `Error rejecting change: ${error instanceof Error ? error.message : String(error)}`,
        changeId
      };
    }
  }
  
  /**
   * Clear all pending changes for a workbook
   * @param workbookId The workbook ID
   */
  public clearPendingChangesForWorkbook(workbookId: string): void {
    const changeIds = this.workbookPendingChanges.get(workbookId) || [];
    
    // Update status of all pending changes to rejected
    changeIds.forEach(id => {
      const change = this.pendingChanges.get(id);
      if (change && change.status === PendingChangeStatus.PENDING) {
        change.status = PendingChangeStatus.REJECTED;
        this.pendingChanges.set(id, change);
      }
    });
    
    // Clear the workbook pending changes mapping
    this.workbookPendingChanges.delete(workbookId);
    
    // Save to storage
    this.saveToStorage();
    
    console.log(`✅ [PendingChangesManager] Cleared all pending changes for workbook: ${workbookId}`);
  }
  
  /**
   * Apply green highlighting to cells affected by pending changes
   * @param workbookId The workbook ID
   */
  public async highlightPendingChanges(workbookId: string): Promise<void> {
    const pendingChanges = this.getPendingChangesForWorkbook(workbookId);
    
    if (pendingChanges.length === 0) {
      console.log(`ℹ️ [PendingChangesManager] No pending changes to highlight for workbook: ${workbookId}`);
      return;
    }
    
    try {
      // Get all affected ranges from pending changes
      const affectedRanges = new Set<string>();
      pendingChanges.forEach(change => {
        change.affectedRanges.forEach(range => affectedRanges.add(range));
      });
      
      // Apply green highlighting to all affected ranges
      await Excel.run(async (context) => {
        for (const rangeAddress of affectedRanges) {
          const { sheet, address } = this.parseReference(rangeAddress);
          const range = context.workbook.worksheets.getItem(sheet).getRange(address);
          range.format.fill.color = "#DDFFDD"; // Light green
        }
        
        await context.sync();
      });
      
      console.log(`✅ [PendingChangesManager] Highlighted ${affectedRanges.size} ranges with pending changes for workbook: ${workbookId}`);
    } catch (error) {
      console.error(`❌ [PendingChangesManager] Error highlighting pending changes for workbook ${workbookId}:`, error);
    }
  }
  
  /**
   * Remove highlighting from cells affected by pending changes
   * @param workbookId The workbook ID
   */
  public async removeHighlighting(workbookId: string): Promise<void> {
    const pendingChanges = this.getPendingChangesForWorkbook(workbookId);
    
    if (pendingChanges.length === 0) {
      console.log(`ℹ️ [PendingChangesManager] No pending changes to remove highlighting for workbook: ${workbookId}`);
      return;
    }
    
    try {
      // Get all affected ranges from pending changes
      const affectedRanges = new Set<string>();
      pendingChanges.forEach(change => {
        change.affectedRanges.forEach(range => affectedRanges.add(range));
      });
      
      // Remove highlighting from all affected ranges
      await Excel.run(async (context) => {
        for (const rangeAddress of affectedRanges) {
          const { sheet, address } = this.parseReference(rangeAddress);
          const range = context.workbook.worksheets.getItem(sheet).getRange(address);
          range.format.fill.clear();
        }
        
        await context.sync();
      });
      
      console.log(`✅ [PendingChangesManager] Removed highlighting from ${affectedRanges.size} ranges for workbook: ${workbookId}`);
    } catch (error) {
      console.error(`❌ [PendingChangesManager] Error removing highlighting for workbook ${workbookId}:`, error);
    }
  }
  
  /**
   * Extract affected ranges from an operation
   * @param operation The Excel operation
   * @returns Array of affected range addresses
   * @private
   */
  private extractAffectedRangesFromOperation(operation: ExcelOperation): string[] {
    const op = operation as any;
    // Get the operation type
    const opType = (op.op || 'unknown');
    const ranges: string[] = [];
    
    switch (opType) {
      case ExcelOperationType.SET_VALUE:
      case ExcelOperationType.ADD_FORMULA:
        if (op.target) {
          ranges.push(op.target);
        }
        break;
        
      case ExcelOperationType.FORMAT_RANGE:
      case ExcelOperationType.CLEAR_RANGE:
      case ExcelOperationType.CREATE_TABLE:
      case ExcelOperationType.SORT_RANGE:
      case ExcelOperationType.FILTER_RANGE:
        if (op.range) {
          ranges.push(op.range);
        }
        break;
        
      case ExcelOperationType.CREATE_SHEET:
      case ExcelOperationType.DELETE_SHEET:
      case ExcelOperationType.RENAME_SHEET:
      case ExcelOperationType.SET_ACTIVE_SHEET:
        // Sheet operations don't have a specific range
        // We could potentially use a special indicator or the first cell
        if (op.name) {
          ranges.push(`${op.name}!A1`);
        } else if (op.sheet) {
          ranges.push(`${op.sheet}!A1`);
        }
        break;
        
      default:
        // Try to extract common range properties
        if (op.target) {
          ranges.push(op.target);
        }
        if (op.range) {
          ranges.push(op.range);
        }
    }
    
    return ranges;
  }
  
  /**
   * Parse a cell reference into sheet and address components
   * @param reference The cell reference (e.g., "Sheet1!A1:B10")
   * @returns Object with sheet and address properties
   * @private
   */
  private parseReference(reference: string): { sheet: string, address: string } {
    const parts = reference.split('!');
    if (parts.length === 2) {
      return {
        sheet: parts[0],
        address: parts[1]
      };
    } else {
      // If no sheet specified, assume the active sheet
      return {
        sheet: '',  // Will be resolved to active sheet
        address: reference
      };
    }
  }
}