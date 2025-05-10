/**
 * PendingChangesTracker
 * 
 * Tracks changes made by AI that require user approval.
 * Highlights affected cells and adds accept/reject buttons on the Excel drawing layer.
 */

import { v4 as uuidv4 } from 'uuid';
import { ExcelOperation } from '../../models/ExcelOperationModels';
import { VersionHistoryService } from '../versioning/VersionHistoryService';
import { ActionRecorder } from '../versioning/ActionRecorder';
import { UndoHandlers } from '../versioning/UndoHandlers';

/**
 * Status of a pending change
 */
export enum PendingChangeStatus {
  PENDING = 'pending',
  ACCEPTED = 'accepted',
  REJECTED = 'rejected'
}

/**
 * Interface for a pending change
 */
export interface PendingChange {
  id: string;
  workbookId: string;
  operation: ExcelOperation;
  affectedRanges: string[];
  status: PendingChangeStatus;
  timestamp: number;
  description: string;
  buttonIds?: {
    accept?: string;
    reject?: string;
  };
}

/**
 * Service for tracking and managing pending changes that require user approval
 */
export class PendingChangesTracker {
  private pendingChanges: Map<string, PendingChange> = new Map();
  private versionHistoryService: VersionHistoryService;
  private actionRecorder: ActionRecorder;
  private undoHandlers: UndoHandlers;
  
  constructor(versionHistoryService: VersionHistoryService, actionRecorder: ActionRecorder) {
    this.versionHistoryService = versionHistoryService;
    this.actionRecorder = actionRecorder;
    this.undoHandlers = new UndoHandlers();
  }
  
  /**
   * Track a new pending change
   * @param workbookId The workbook ID
   * @param operation The Excel operation
   * @returns The created pending change
   */
  public trackChange(workbookId: string, operation: ExcelOperation): PendingChange {
    const id = uuidv4();
    const affectedRanges = this.extractAffectedRanges(operation);
    
    const pendingChange: PendingChange = {
      id,
      workbookId,
      operation,
      affectedRanges,
      status: PendingChangeStatus.PENDING,
      timestamp: Date.now(),
      description: this.generateChangeDescription(operation)
    };
    
    this.pendingChanges.set(id, pendingChange);
    
    // Highlight the affected cells and add approval buttons
    this.highlightChanges(pendingChange);
    
    return pendingChange;
  }
  
  /**
   * Accept a pending change
   * @param changeId The ID of the pending change to accept
   */
  public async acceptChange(changeId: string): Promise<void> {
    const change = this.pendingChanges.get(changeId);
    if (!change) {
      console.error(`Cannot accept change: Change with ID ${changeId} not found`);
      return;
    }
    
    if (change.status !== PendingChangeStatus.PENDING) {
      console.warn(`Change with ID ${changeId} is already ${change.status}`);
      return;
    }
    
    // Update status
    change.status = PendingChangeStatus.ACCEPTED;
    this.pendingChanges.set(changeId, change);
    
    // Remove highlighting
    await this.removeHighlighting(change);
    
    console.log(`✅ Accepted change: ${change.description}`);
  }
  
  /**
   * Reject a pending change
   * @param changeId The ID of the pending change to reject
   */
  public async rejectChange(changeId: string): Promise<void> {
    const change = this.pendingChanges.get(changeId);
    if (!change) {
      console.error(`Cannot reject change: Change with ID ${changeId} not found`);
      return;
    }
    
    if (change.status !== PendingChangeStatus.PENDING) {
      console.warn(`Change with ID ${changeId} is already ${change.status}`);
      return;
    }
    
    // Update status
    change.status = PendingChangeStatus.REJECTED;
    this.pendingChanges.set(changeId, change);
    
    // Undo the operation
    await this.undoOperation(change);
    
    // Remove highlighting
    await this.removeHighlighting(change);
    
    console.log(`❌ Rejected change: ${change.description}`);
  }
  
  /**
   * Get all pending changes for a workbook
   * @param workbookId The workbook ID
   * @returns Array of pending changes
   */
  public getPendingChanges(workbookId: string): PendingChange[] {
    return Array.from(this.pendingChanges.values())
      .filter(change => change.workbookId === workbookId && change.status === PendingChangeStatus.PENDING);
  }
  
  /**
   * Highlight cells affected by a pending change
   * @param change The pending change to highlight
   * @private
   */
  private async highlightChanges(change: PendingChange): Promise<void> {
    if (change.affectedRanges.length === 0) {
      console.warn(`No affected ranges found for change: ${change.id}`);
      return;
    }
    
    try {
      await Excel.run(async (context) => {
        for (const rangeAddress of change.affectedRanges) {
          // Parse the range address
          const { sheet, address } = this.parseReference(rangeAddress);
          
          // Get the worksheet and range
          const worksheet = context.workbook.worksheets.getItem(sheet);
          const range = worksheet.getRange(address);
          
          // Apply light green highlight to the range
          range.format.fill.color = '#DDFFDD';
        }
        
        await context.sync();
      });
      
      console.log(`✅ Highlighted changes for: ${change.description}`);
    } catch (error) {
      console.error(`Error highlighting changes: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  
  /**
   * Remove highlighting for a pending change
   * @param change The pending change
   * @private
   */
  private async removeHighlighting(change: PendingChange): Promise<void> {
    try {
      await Excel.run(async (context) => {
        // Remove highlighting from affected ranges
        for (const rangeAddress of change.affectedRanges) {
          const { sheet, address } = this.parseReference(rangeAddress);
          const worksheet = context.workbook.worksheets.getItem(sheet);
          const range = worksheet.getRange(address);
          
          // Only clear the fill if it's our light green color
          range.format.load('fill/color');
          await context.sync();
          
          if (range.format.fill.color === '#DDFFDD') {
            range.format.fill.clear();
          }
        }
        
        await context.sync();
      });
      
      console.log(`✅ Removed highlighting for: ${change.description}`);
    } catch (error) {
      console.error(`Error removing highlighting: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  
  /**
   * Undo an operation
   * @param change The pending change to undo
   * @private
   */
  private async undoOperation(change: PendingChange): Promise<void> {
    try {
      await Excel.run(async (context) => {
        // The implementation will depend on the operation type
        // For now, we'll use a simple approach based on the operation type
        
        const op = change.operation;
        const opType = op.op;
        
        switch (opType) {
          case 'set_value':
            // Revert value changes
            const setValueOp = op as any;
            if (setValueOp.target) {
              const { sheet, address } = this.parseReference(setValueOp.target);
              const range = context.workbook.worksheets.getItem(sheet).getRange(address);
              range.clear();
            }
            break;
            
          case 'format_range':
            // Revert formatting changes
            const formatOp = op as any;
            if (formatOp.range) {
              const { sheet, address } = this.parseReference(formatOp.range);
              const range = context.workbook.worksheets.getItem(sheet).getRange(address);
              
              // Clear specific formatting based on what was changed
              if (formatOp.style) {
                range.numberFormat = [['General']];
              }
              if (formatOp.bold !== undefined) {
                range.format.font.bold = false;
              }
              if (formatOp.italic !== undefined) {
                range.format.font.italic = false;
              }
              if (formatOp.fontColor) {
                range.format.font.color = 'Automatic';
              }
              if (formatOp.fillColor) {
                range.format.fill.clear();
              }
            }
            break;
            
          // Add more cases for other operation types as needed
          
          default:
            console.warn(`Undo not implemented for operation type: ${opType}`);
            // For operations without specific undo logic, try to clear the affected ranges
            for (const rangeAddress of change.affectedRanges) {
              const { sheet, address } = this.parseReference(rangeAddress);
              const range = context.workbook.worksheets.getItem(sheet).getRange(address);
              range.clear();
            }
        }
        
        await context.sync();
      });
      
      console.log(`✅ Undid operation: ${change.description}`);
    } catch (error) {
      console.error(`Error undoing operation: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  
  /**
   * Extract affected ranges from an operation
   * @param operation The Excel operation
   * @returns Array of affected range addresses
   * @private
   */
  private extractAffectedRanges(operation: ExcelOperation): string[] {
    const op = operation as any;
    const opType = op.op || 'unknown';
    const ranges: string[] = [];
    
    switch (opType) {
      case 'set_value':
      case 'add_formula':
        if (op.target) {
          ranges.push(op.target);
        }
        break;
        
      case 'format_range':
      case 'clear_range':
      case 'create_table':
      case 'sort_range':
      case 'filter_range':
        if (op.range) {
          ranges.push(op.range);
        }
        break;
        
      case 'create_sheet':
      case 'delete_sheet':
      case 'rename_sheet':
      case 'set_active_sheet':
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
   * Generate a description for a change based on the operation
   * @param operation The Excel operation
   * @returns A human-readable description
   * @private
   */
  private generateChangeDescription(operation: ExcelOperation): string {
    const op = operation as any;
    const opType = op.op || 'unknown';
    
    switch (opType) {
      case 'set_value':
        return `Set value in ${op.target}`;
        
      case 'add_formula':
        return `Add formula to ${op.target}`;
        
      case 'format_range':
        return `Format range ${op.range}`;
        
      case 'create_sheet':
        return `Create sheet "${op.name}"`;
        
      case 'delete_sheet':
        return `Delete sheet "${op.name}"`;
        
      case 'rename_sheet':
        return `Rename sheet from "${op.sheet}" to "${op.name}"`;
        
      case 'set_active_sheet':
        return `Set active sheet to "${op.name}"`;
        
      case 'clear_range':
        return `Clear range ${op.range}`;
        
      case 'create_table':
        return `Create table in range ${op.range}`;
        
      case 'sort_range':
        return `Sort range ${op.range}`;
        
      case 'filter_range':
        return `Filter range ${op.range}`;
        
      default:
        return `${opType} operation`;
    }
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
  
  /**
   * Set up event handlers for the accept/reject buttons
   * This needs to be called once to register the handlers
   */
  public setupButtonEventHandlers(): void {
    // We need to use Office.js events to handle shape clicks
    // This is a placeholder for the implementation
    console.log('Setting up button event handlers');
    
    // In a real implementation, we would set up event handlers for shape clicks
    // and map them to the appropriate accept/reject functions
  }
  
  /**
   * Refresh the highlighting for all pending changes
   * This should be called periodically to ensure the highlighting remains visible
   * @param workbookId The workbook ID
   */
  public async refreshPendingChangesHighlighting(workbookId: string): Promise<void> {
    const pendingChanges = this.getPendingChanges(workbookId);
    
    if (pendingChanges.length === 0) {
      return;
    }
    
    console.log(`🔄 [PendingChangesTracker] Refreshing highlighting for ${pendingChanges.length} pending changes`);
    
    try {
      await Excel.run(async (context) => {
        for (const change of pendingChanges) {
          // Reapply the highlighting
          await this.highlightChanges(change);
        }
        
        await context.sync();
      });
      
      console.log(`✅ [PendingChangesTracker] Successfully refreshed highlighting for pending changes`);
    } catch (error) {
      console.error(`❌ [PendingChangesTracker] Error refreshing highlighting: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}
