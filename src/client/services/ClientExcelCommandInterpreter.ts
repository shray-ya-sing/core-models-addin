// src/client/services/ClientExcelCommandInterpreter.ts
// Interprets and executes Excel operations using the Office.js API

import { 
  ExcelOperation, 
  ExcelOperationType,
  ExcelCommandPlan,
  CreateScenarioTableOperation,
  SetRowColumnOptionsOperation,
  SetCalculationOptionsOperation,
  RecalculateRangesOperation
} from '../models/ExcelOperationModels';
import * as ExcelUtils from '../utils/ExcelUtils';
import { ActionRecorder } from './versioning/ActionRecorder';
import { VersionHistoryService } from './versioning/VersionHistoryService';
import { PendingChangesTracker } from './PendingChangesTracker';
import { ShapeEventHandler } from './ShapeEventHandler';

/**
 * Service that interprets and executes Excel operations using Office.js
 */
export class ClientExcelCommandInterpreter {
  private actionRecorder: ActionRecorder | null = null;
  private pendingChangesTracker: PendingChangesTracker | null = null;
  private shapeEventHandler: ShapeEventHandler | null = null;
  private currentWorkbookId: string = '';
  private requireApproval: boolean = false;
  
  /**
   * Set the action recorder for version history tracking
   * @param actionRecorder The action recorder instance
   */
  public setActionRecorder(actionRecorder: ActionRecorder): void {
    console.log(`🔄 [ClientExcelCommandInterpreter] Setting ActionRecorder instance`);
    this.actionRecorder = actionRecorder;
  }
  
  /**
   * Set the pending changes tracker for AI-generated changes approval
   * @param pendingChangesTracker The pending changes tracker instance
   * @param shapeEventHandler The shape event handler instance
   */
  public setPendingChangesTracker(pendingChangesTracker: PendingChangesTracker, shapeEventHandler: ShapeEventHandler): void {
    console.log(`🔄 [ClientExcelCommandInterpreter] Setting PendingChangesTracker instance`);
    this.pendingChangesTracker = pendingChangesTracker;
    this.shapeEventHandler = shapeEventHandler;
    
    // Start polling for shape selection events
    this.shapeEventHandler.startPolling();
  }
  
  /**
   * Enable or disable the approval workflow for AI-generated changes
   * @param enable Whether to enable the approval workflow
   */
  public setRequireApproval(enable: boolean): void {
    console.log(`🔄 [ClientExcelCommandInterpreter] ${enable ? 'Enabling' : 'Disabling'} approval workflow for AI-generated changes`);
    this.requireApproval = enable;
  }
  
  /**
   * Get the current action recorder instance
   * @returns The current action recorder instance or null if not set
   */
  public getActionRecorder(): ActionRecorder | null {
    return this.actionRecorder;
  }
  
  /**
   * Set the current workbook ID for version history tracking
   * @param workbookId The current workbook ID
   */
  public setCurrentWorkbookId(workbookId: string): void {
    console.log(`🔄 [ClientExcelCommandInterpreter] Setting current workbook ID: ${workbookId}`);
    this.currentWorkbookId = workbookId;
  }
  
  /**
   * Get the current workbook ID
   * @returns The current workbook ID or empty string if not set
   */
  public getCurrentWorkbookId(): string {
    return this.currentWorkbookId;
  }
  
  /**
   * Creates a unique fingerprint for an operation based on its properties
   * This helps detect duplicate operations even if they have different IDs
   * @param operation The Excel operation to create a fingerprint for
   * @returns A string fingerprint that uniquely identifies this operation
   * @private
   */
  private createOperationFingerprint(operation: ExcelOperation): string {
    // Extract key properties that identify an operation
    const opType = operation.op || 'unknown';
    
    // Safely extract properties based on operation type
    let targetOrRange = '';
    let valueOrFormula = '';
    let formatInfo = '';
    
    // Get properties safely using type assertion
    const op = operation as any;
    
    // Extract common properties based on operation type string
    if (opType === 'set_value') {
      targetOrRange = op.target || '';
      valueOrFormula = op.value !== undefined ? String(op.value) : '';
    } 
    else if (opType === 'add_formula') {
      targetOrRange = op.target || '';
      valueOrFormula = op.formula || '';
    }
    else if (opType === 'format_range') {
      targetOrRange = op.range || '';
      formatInfo = op.style || '';
    }
    else if (opType === 'create_sheet') {
      valueOrFormula = op.name || '';
    }
    else if (['delete_sheet', 'rename_sheet', 'set_active_sheet'].includes(opType)) {
      targetOrRange = op.name || op.sheet || '';
    }
    else {
      // For other operations, try to get common properties
      targetOrRange = op.target || op.range || op.sheet || '';
      valueOrFormula = op.value || op.formula || op.name || '';
    }
    
    // Create a fingerprint string that combines these properties
    const fingerprint = `${opType}:${targetOrRange}:${valueOrFormula}:${formatInfo}`;
    
    return fingerprint;
  }
  
  /**
   * Execute a command plan with multiple operations
   * @param plan The Excel command plan to execute
   * @returns A promise that resolves when all operations are complete
   */
  public async executeCommandPlan(plan: ExcelCommandPlan): Promise<void> {
    console.log(`Executing command plan: ${plan.description}`);
    console.log(`Operations to execute: ${plan.operations.length}`);
    
    try {
      // Check if we have operations with dependencies
      const hasDependencies = plan.operations.some(op => op.dependsOn && op.dependsOn.length > 0);
      
      if (hasDependencies) {
        return this.executeOperationsWithDependencies(plan.operations);
      } else {
        return this.executeOperations(plan.operations);
      }
    } catch (error) {
      // Record the error in the plan if it has an ID
      if (plan.id) {
        plan.error = {
          message: error.message,
          details: error
        };
      }
      throw error;
    }
  }
  
  /**
   * Execute a list of Excel operations
   * @param operations The operations to execute
   * @returns A promise that resolves when all operations are complete
   */
  public async executeOperations(operations: ExcelOperation[]): Promise<void> {
    if (!operations || operations.length === 0) {
      console.warn('No operations to execute');
      return;
    }
    
    // Create a unique ID for this batch of operations for tracking
    const batchId = Math.random().toString(36).substring(2, 10);
    
    // Log details about the operations being executed
    console.log(`📣 [ClientExcelCommandInterpreter] executeOperations called with ${operations.length} operations (batch: ${batchId})`);
    
    // Track operation IDs to detect duplicates within this batch
    const operationIds = new Set<string>();
    const operationsByType = new Map<string, number>();
    
    // Log each operation
    operations.forEach((op, index) => {
      const opType = op.op || 'unknown';
      const opId = op.id || 'no-id';
      
      // Count operations by type
      const typeCount = operationsByType.get(opType) || 0;
      operationsByType.set(opType, typeCount + 1);
      
      // Check for duplicate IDs within this batch
      if (opId !== 'no-id') {
        if (operationIds.has(opId)) {
          console.warn(`⚠️ [ClientExcelCommandInterpreter] Duplicate operation ID ${opId} within batch ${batchId}`);
        } else {
          operationIds.add(opId);
        }
      }
      
      // Log operation details for the first 5 operations (to avoid excessive logging)
      if (index < 5) {
        console.log(`📝 [ClientExcelCommandInterpreter] Operation ${index} in batch ${batchId}: ${opType} (ID: ${opId})`);
      }
    });
    
    // Log summary of operation types
    console.log(`📊 [ClientExcelCommandInterpreter] Operations by type in batch ${batchId}:`, 
      Array.from(operationsByType.entries())
        .map(([type, count]) => `${type}: ${count}`)
        .join(', '));
    
    try {
      await Excel.run(async (context) => {
        for (const operation of operations) {
          await this.executeOperation(context, operation);
        }
        
        await context.sync();
      });
      
      console.log(`✅ [ClientExcelCommandInterpreter] All operations in batch ${batchId} executed successfully`);
    } catch (error) {
      console.error(`❌ [ClientExcelCommandInterpreter] Error executing Excel operations in batch ${batchId}:`, error);
      throw error;
    }
  }
  
  /**
   * Execute a single Excel operation
   * @param context The Excel context
   * @param operation The operation to execute
   */
  /**
   * Execute operations with dependency handling
   * @param operations The operations to execute
   * @returns A promise that resolves when all operations are complete
   */
  public async executeOperationsWithDependencies(operations: ExcelOperation[]): Promise<void> {
    if (!operations || operations.length === 0) {
      console.warn('No operations to execute');
      return;
    }
    
    // Build dependency graph
    const sortedOperations = this.sortOperationsByDependency(operations);
    
    try {
      await Excel.run(async (context) => {
        for (const operation of sortedOperations) {
          try {
            await this.executeOperation(context, operation);
            if (operation.id) {
              console.log(`Successfully executed operation: ${operation.id}`);
            }
          } catch (error) {
            console.error(`Error executing operation ${operation.id || operation.op}:`, error);
            
            // Skip to the next operation if this one should ignore errors
            if (operation.ignoreErrors) {
              console.warn(`Ignoring error and continuing execution`);
              continue;
            }
            
            // Otherwise re-throw the error
            throw error;
          }
        }
        
        await context.sync();
      });
      
      console.log('All operations executed successfully');
    } catch (error) {
      console.error('Error executing Excel operations:', error);
      throw error;
    }
  }
  
  /**
   * Sort operations by dependency to ensure they execute in the correct order
   * @param operations The operations to sort
   * @returns Sorted operations
   */
  private sortOperationsByDependency(operations: ExcelOperation[]): ExcelOperation[] {
    // Create map of operation IDs to operations
    const operationMap = new Map<string, ExcelOperation>();
    operations.forEach(op => {
      if (op.id) {
        operationMap.set(op.id, op);
      }
    });
    
    // Create a graph of dependencies
    const graph = new Map<string, string[]>();
    const noIds: ExcelOperation[] = [];
    
    operations.forEach(op => {
      if (!op.id) {
        // Operations without IDs have no dependencies and go first
        noIds.push(op);
        return;
      }
      
      graph.set(op.id, op.dependsOn || []);
    });
    
    // Perform topological sort
    const visited = new Set<string>();
    const temp = new Set<string>();
    const result: ExcelOperation[] = [];
    
    // Start with operations that have no dependencies (no IDs)
    result.push(...noIds);
    
    // Function to visit a node in the dependency graph
    const visit = (id: string) => {
      if (!operationMap.has(id)) {
        console.warn(`Operation with ID ${id} not found but is referenced as a dependency`);
        return;
      }
      
      // If we've already processed this operation, skip it
      if (visited.has(id)) return;
      
      // Check for circular dependencies
      if (temp.has(id)) {
        throw new Error(`Circular dependency detected involving operation ${id}`);
      }
      
      // Mark as being processed
      temp.add(id);
      
      // Process all dependencies first
      const dependencies = graph.get(id) || [];
      dependencies.forEach(depId => visit(depId));
      
      // Mark as processed and add to result
      temp.delete(id);
      visited.add(id);
      result.push(operationMap.get(id)!);
    };
    
    // Visit each operation
    operations.forEach(op => {
      if (op.id && !visited.has(op.id)) {
        visit(op.id);
      }
    });
    
    return result;
  }

  private async executeOperation(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    // Generate a unique execution ID for this operation execution
    const executionId = Math.random().toString(36).substring(2, 10);
    const opId = operation.id || 'no-id';
    const opType = operation.op || 'unknown';
    
    // Create a fingerprint for this operation
    const fingerprint = this.createOperationFingerprint(operation);
    
    // Record the operation for version history if action recorder is available
    console.log(`🔄 [ClientExcelCommandInterpreter] Executing operation: ${opType} (ID: ${opId}, Execution: ${executionId}, Fingerprint: ${fingerprint})`);
    
    // Variable to track if this operation requires approval
    let requiresApproval = false;
    let pendingChangeId = '';
    
    // Check if this operation requires approval
    if (this.requireApproval && this.pendingChangesTracker && this.currentWorkbookId) {
      console.log(`🔍 [ClientExcelCommandInterpreter] Operation requires approval: ${opType} (ID: ${opId})`);
      
      // Track the operation as a pending change
      const pendingChange = this.pendingChangesTracker.trackChange(this.currentWorkbookId, operation);
      pendingChangeId = pendingChange.id;
      requiresApproval = true;
      console.log(`✅ [ClientExcelCommandInterpreter] Added operation to pending changes: ${pendingChange.id}`);
      
      // Continue with execution to show the changes to the user
      // The changes will be highlighted to indicate they are pending approval
    }
    
    // If no approval required or approval system not set up, proceed with normal execution
    if (this.actionRecorder && this.currentWorkbookId) {
      console.log(`📝 [ClientExcelCommandInterpreter] Recording operation ${opType} for workbook: ${this.currentWorkbookId} (Execution: ${executionId})`);
      try {
        // Record the operation before executing it
        await this.actionRecorder.recordOperation(context, this.currentWorkbookId, operation);
        console.log(`✅ [ClientExcelCommandInterpreter] Successfully recorded operation ${opType} (Execution: ${executionId})`);
      } catch (error) {
        // Log error but continue with operation execution
        console.error(`❌ [ClientExcelCommandInterpreter] Error recording operation for version history (Execution: ${executionId}):`, error);
      }
    } else {
      if (!this.actionRecorder) {
        console.warn(`⚠️ [ClientExcelCommandInterpreter] Cannot record operation: ActionRecorder not set (Execution: ${executionId})`);
      }
      if (!this.currentWorkbookId) {
        console.warn(`⚠️ [ClientExcelCommandInterpreter] Cannot record operation: workbookId not set (Execution: ${executionId})`);
      }
    }
    
    try {
      // Execute the operation based on its type
      switch (operation.op) {
        case ExcelOperationType.SET_VALUE:
          await this.executeSetValue(context, operation);
          break;
          
        case ExcelOperationType.ADD_FORMULA:
          await this.executeAddFormula(context, operation);
          break;
          
        case ExcelOperationType.CREATE_CHART:
          await this.executeCreateChart(context, operation);
          break;
          
        case ExcelOperationType.FORMAT_RANGE:
          await this.executeFormatRange(context, operation);
          break;
          
        case ExcelOperationType.CLEAR_RANGE:
          await this.executeClearRange(context, operation);
          break;
          
        case ExcelOperationType.CREATE_TABLE:
          await this.executeCreateTable(context, operation);
          break;
          
        case ExcelOperationType.SORT_RANGE:
          await this.executeSortRange(context, operation);
          break;
          
        case ExcelOperationType.FILTER_RANGE:
          await this.executeFilterRange(context, operation);
          break;
          
        case ExcelOperationType.CREATE_SHEET:
          await this.executeCreateSheet(context, operation);
          break;
          
        case ExcelOperationType.DELETE_SHEET:
          await this.executeDeleteSheet(context, operation);
          break;
          
          
        case ExcelOperationType.COPY_RANGE:
          await this.executeCopyRange(context, operation);
          break;
          
        case ExcelOperationType.MERGE_CELLS:
          await this.executeMergeCells(context, operation);
          break;
          
        case ExcelOperationType.UNMERGE_CELLS:
          await this.executeUnmergeCells(context, operation);
          break;
          
        case ExcelOperationType.CONDITIONAL_FORMAT:
          await this.executeConditionalFormat(context, operation);
          break;
          
        case ExcelOperationType.ADD_COMMENT:
          await this.executeAddComment(context, operation);
          break;
          
        case ExcelOperationType.SET_FREEZE_PANES:
          await this.executeSetFreezePanes(context, operation);
          break;

          
        case ExcelOperationType.SET_ACTIVE_SHEET:
          await this.executeSetActiveSheet(context, operation);
          break;

        case ExcelOperationType.SET_WORKSHEET_SETTINGS:
          await this.executeSetWorksheetSettings(context, operation);
          break;


        // Print and page layout operations
        case ExcelOperationType.SET_PRINT_SETTINGS:
          await this.executeSetPrintSettings(context, operation);
          break;
          
        case ExcelOperationType.SET_PAGE_SETUP:
          await this.executeSetPageSetup(context, operation);
          break;
          
        // Chart formatting operations
        case ExcelOperationType.FORMAT_CHART:
          await this.executeFormatChart(context, operation);
          break;          
          
        // Complex operations
        case ExcelOperationType.COMPOSITE_OPERATION:
          await this.executeCompositeOperation(context, operation);
          break;
          
        case ExcelOperationType.BATCH_OPERATION:
          await this.executeBatchOperation(context, operation);
          break;
          
        case ExcelOperationType.EXPORT_TO_PDF:
          await this.executeExportToPdf(context, operation);
          break;
          
        case ExcelOperationType.CREATE_SCENARIO_TABLE:
          await this.executeCreateScenarioTable(context, operation);
          break;
          
        case ExcelOperationType.SET_ROW_COLUMN_OPTIONS:
          await this.executeSetRowColumnOptions(context, operation);
          break;

        case ExcelOperationType.SET_CALCULATION_OPTIONS:
          await this.executeSetCalculationOptions(context, operation);
          break;

        case ExcelOperationType.RECALCULATE_RANGES:
          await this.executeRecalculateRanges(context, operation);
          break;
          
        default:
          console.warn(`Unsupported operation: ${(operation as any).op}`);
      }
      
      // Apply highlighting if this operation requires approval
      if (requiresApproval && this.pendingChangesTracker && pendingChangeId && this.currentWorkbookId) {
        console.log(`🎨 [ClientExcelCommandInterpreter] Applying highlighting for pending change: ${pendingChangeId}`);
        try {
          // Since we just created the pending change, we know it exists
          // We'll apply highlighting to all pending changes for the current workbook
          // This ensures the cells are properly highlighted
          await Excel.run(async (context) => {
            // Extract affected ranges from the operation
            const affectedRanges = this.getAffectedRanges(operation);
            
            // Apply green highlighting to all affected ranges
            for (const rangeAddress of affectedRanges) {
              const { sheet, address } = this.parseReference(rangeAddress);
              const range = context.workbook.worksheets.getItem(sheet).getRange(address);
              range.format.fill.color = "#DDFFDD"; // Light green
            }
            
            await context.sync();
          });
          
          console.log(`✅ [ClientExcelCommandInterpreter] Successfully highlighted affected ranges for pending change: ${pendingChangeId}`);
        } catch (highlightError) {
          // Log error but don't fail the operation
          console.error(`❌ [ClientExcelCommandInterpreter] Error highlighting pending change: ${highlightError instanceof Error ? highlightError.message : String(highlightError)}`);
        }
      }
    } catch (error) {
      console.error(`Error executing operation ${operation.op}:`, error);
      throw error;
    }
  }
  
  /**
   * Parse a cell reference into sheet name and address
   * @param reference The cell reference (e.g. "Sheet1!A1")
   * @returns An object with sheet name and address
   */
  private parseReference(reference: string): { sheet: string; address: string } {
    const parts = reference.split('!');
    if (parts.length !== 2) {
      throw new Error(`Invalid cell reference: ${reference}`);
    }
    
    return {
      sheet: parts[0],
      address: parts[1]
    };
  }
  
  /**
   * Extract affected ranges from an operation
   * @param operation The Excel operation
   * @returns Array of affected range addresses (e.g., "Sheet1!A1:B10")
   */
  private getAffectedRanges(operation: ExcelOperation): string[] {
    const affectedRanges: string[] = [];
    const op = operation as any;
    
    switch (operation.op) {
      case ExcelOperationType.SET_VALUE:
      case ExcelOperationType.ADD_FORMULA:
        if (op.target) {
          affectedRanges.push(op.target);
        }
        break;
        
      case ExcelOperationType.FORMAT_RANGE:
      case ExcelOperationType.CLEAR_RANGE:
      case ExcelOperationType.SORT_RANGE:
      case ExcelOperationType.FILTER_RANGE:
      case ExcelOperationType.MERGE_CELLS:
      case ExcelOperationType.UNMERGE_CELLS:
      case ExcelOperationType.CONDITIONAL_FORMAT:
        if (op.range) {
          affectedRanges.push(op.range);
        }
        break;
        
      case ExcelOperationType.COPY_RANGE:
        if (op.source) {
          affectedRanges.push(op.source);
        }
        if (op.destination) {
          affectedRanges.push(op.destination);
        }
        break;
        
      case ExcelOperationType.CREATE_TABLE:
        if (op.range) {
          affectedRanges.push(op.range);
        }
        break;
        
      case ExcelOperationType.CREATE_CHART:
        if (op.dataRange) {
          affectedRanges.push(op.dataRange);
        }
        break;
        
      // For sheet-level operations, we don't have specific ranges to highlight
      // For other operations, try to extract any range-like properties
      default:
        // Try to find any properties that might contain range references
        for (const key of ['range', 'target', 'source', 'destination', 'dataRange']) {
          if (op[key] && typeof op[key] === 'string' && op[key].includes('!')) {
            affectedRanges.push(op[key]);
          }
        }
    }
    
    return affectedRanges;
  }
  
  /**
   * Execute a SET_VALUE operation
   */
  private async executeSetValue(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.target);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.values = [[op.value]];
  }
  
  /**
   * Execute an ADD_FORMULA operation
   */
  private async executeAddFormula(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.target);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.formulas = [[op.formula]];
  }
  
  /**
   * Execute a CREATE_CHART operation
   */
  private async executeCreateChart(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const worksheet = context.workbook.worksheets.getItem(sheet);
    const range = worksheet.getRange(address);
    
    const chart = worksheet.charts.add(op.type, range, Excel.ChartSeriesBy.auto);
    
    if (op.title) {
      chart.title.text = op.title;
    }
    
    if (op.position) {
      const { sheet: posSheet, address: posAddress } = this.parseReference(op.position);
      if (posSheet === sheet) {
        const positionRange = worksheet.getRange(posAddress);
        chart.setPosition(positionRange);
      }
    }
  }
  
  /**
   * Execute a FORMAT_RANGE operation
   */
  private async executeFormatRange(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    const format = range.format;
    
    if (op.style) {
      // Apply number format - numberFormat requires a 2D array
      let formatString: string;
      
      switch (op.style.toLowerCase()) {
        case 'currency':
          formatString = '"$"#,##0.00';
          break;
        case 'percentage_2_decimal':
          formatString = '0.00%';
          break;
        case 'percentage_1_decimal':
          formatString = '0.0%';
          break;
        case 'percentage_0_decimal':
          formatString = '0%';
          break;
        case 'date':
          formatString = 'm/d/yyyy';
          break;
        case 'time':
          formatString = 'h:mm:ss AM/PM';
          break;
        case 'scientific':
          formatString = '0.00E+00';
          break;
        case 'text':
          formatString = '@';
          break;
        case 'financial':
          formatString = '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)';
          break;
        case 'financial_with_dollar':
          formatString = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)';
          break;
        case 'financial_with_euro':
          formatString = '_(€* #,##0_);_(€* (#,##0);_(€* "-"??_);_(@_)';
          break;
        case 'financial_with_pound':
          formatString = '_(£* #,##0_);_(£* (#,##0);_(£* "-"??_);_(@_)';
          break;
        case 'financial_with_yen':
          formatString = '_(¥* #,##0_);_(¥* (#,##0);_(¥* "-"??_);_(@_)';
          break;
        case 'financial_with_rupee':
          formatString = '_(₹* #,##0_);_(₹* (#,##0);_(₹* "-"??_);_(@_)';
          break;
        case 'financial_with_yen':
          formatString = '_(¥* #,##0_);_(¥* (#,##0);_(¥* "-"??_);_(@_)';
          break;
        case 'projection_year':
          formatString = '#"E"';
          break;
        case 'actual_year':
          formatString = '#"A"';
          break;
        case 'number':
          formatString = '0.00';
          break;
        default:
          formatString = op.style;
      }

      if(op.customNumberFormat){
        formatString = op.customNumberFormat;
      }
      
      // Create a 2D array with the same format for all cells in the range
      // We'll use a single-cell format and Excel will apply it to the entire range
      range.numberFormat = [[formatString]];
    }
    
    if (op.bold !== undefined) {
      format.font.bold = op.bold;
    }
    
    if (op.italic !== undefined) {
      format.font.italic = op.italic;
    }
    
    if (op.fontColor) {
      format.font.color = op.fontColor;
    }
    
    if (op.fillColor) {
      format.fill.color = op.fillColor;
    }
    
    if (op.fontSize) {
      format.font.size = op.fontSize;
    }
    
    if (op.horizontalAlignment) {
      switch (op.horizontalAlignment.toLowerCase()) {
        case 'left':
          format.horizontalAlignment = Excel.HorizontalAlignment.left;
          break;
        case 'center':
          format.horizontalAlignment = Excel.HorizontalAlignment.center;
          break;
        case 'right':
          format.horizontalAlignment = Excel.HorizontalAlignment.right;
          break;
        case 'justify':
          format.horizontalAlignment = Excel.HorizontalAlignment.justify;
          break;
        case 'distributed':
          format.horizontalAlignment = Excel.HorizontalAlignment.distributed;
          break;
        case 'center_across_selection':
          format.horizontalAlignment = Excel.HorizontalAlignment.centerAcrossSelection;
          break;
      }
    }
    
    if (op.verticalAlignment) {
      switch (op.verticalAlignment.toLowerCase()) {
        case 'top':
          format.verticalAlignment = Excel.VerticalAlignment.top;
          break;
        case 'center':
          format.verticalAlignment = Excel.VerticalAlignment.center;
          break;
        case 'bottom':
          format.verticalAlignment = Excel.VerticalAlignment.bottom;
          break;
        case 'justify':
          format.verticalAlignment = Excel.VerticalAlignment.justify;
          break;
        case 'distributed':
          format.verticalAlignment = Excel.VerticalAlignment.distributed;
          break;
      }
    }

    if(op.indent){
      range.format.indentLevel = op.indent;
    }

    if(op.wrapText){
      range.format.wrapText = op.wrapText;
    }

    if(op.shrinkToFit){
      range.format.shrinkToFit = op.shrinkToFit;
    }

    if(op.textOrientation){
      range.format.textOrientation = op.textOrientation;
    }

  }
  
  /**
   * Execute a CLEAR_RANGE operation
   */
  private async executeClearRange(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    
    if (!op.clearType || op.clearType === 'all') {
      range.clear(Excel.ClearApplyTo.all);
    } else if (op.clearType === 'formats') {
      range.clear(Excel.ClearApplyTo.formats);
    } else if (op.clearType === 'contents') {
      range.clear(Excel.ClearApplyTo.contents);
    }
  }
  
  /**
   * Execute a CREATE_TABLE operation
   */
  private async executeCreateTable(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    const table = range.worksheet.tables.add(range, op.hasHeaders ?? true);
    
    if (op.styleName) {
      table.style = op.styleName;
    }
  }
  
  /**
   * Execute a SORT_RANGE operation
   */
  private async executeSortRange(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    
    // Convert column letter to index (0-based)
    const columnIndex = op.sortBy.charCodeAt(0) - 'A'.charCodeAt(0);
    
    // Create a sort field
    const sortField: Excel.SortField = {
      key: columnIndex,
      ascending: op.sortDirection.toLowerCase() === 'ascending'
    };
    
    // Apply the sort
    range.sort.apply([sortField], op.hasHeaders ?? true);
  }
  
  /**
   * Execute a FILTER_RANGE operation
   */
  private async executeFilterRange(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    
    // Convert column letter to index (0-based)
    const columnIndex = op.column.charCodeAt(0) - 'A'.charCodeAt(0);
    
    // Create a table if not already one
    const table = range.worksheet.tables.add(range, true);
    
    // Apply the filter
    table.columns.getItem(columnIndex).filter.applyCustomFilter(op.criteria);
  }
  
  /**
   * Execute a CREATE_SHEET operation
   */
  private async executeCreateSheet(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    context.workbook.worksheets.add(op.name);
  }
  
  /**
   * Execute a DELETE_SHEET operation
   */
  private async executeDeleteSheet(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    context.workbook.worksheets.getItem(op.name).delete();
  }
  

  
  /**
   * Execute a COPY_RANGE operation
   */
  private async executeCopyRange(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet: sourceSheet, address: sourceAddress } = this.parseReference(op.source);
    const { sheet: destSheet, address: destAddress } = this.parseReference(op.destination);
    
    const sourceRange = context.workbook.worksheets.getItem(sourceSheet).getRange(sourceAddress);
    const destRange = context.workbook.worksheets.getItem(destSheet).getRange(destAddress);
    
    // Copy values, formulas, and formats
    sourceRange.copyFrom(
      sourceRange,
      Excel.RangeCopyType.all,
      false, // skipBlanks
      false  // transpose
    );
    
  }
  
  /**
   * Execute a MERGE_CELLS operation
   */
  private async executeMergeCells(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.merge();
  }
  
  /**
   * Execute an UNMERGE_CELLS operation
   */
  private async executeUnmergeCells(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.unmerge();
  }
  
  /**
   * Execute a CONDITIONAL_FORMAT operation
   */
  private async executeConditionalFormat(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.range);
    
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    const conditionalFormats = range.conditionalFormats;
    
    switch (op.type.toLowerCase()) {
      case 'databar':
        conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        break;
        
      case 'colorscale':
        conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        break;
        
      case 'iconset':
        conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
        break;
        
      case 'topbottom':
        const topBottom = conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
        // Default to top 10 items if no rank is specified
        const rank = op.value ? Number(op.value) : 10;
        
        topBottom.topBottom.rule = {
          type: op.criteria?.toLowerCase().includes('top') 
            ? Excel.ConditionalTopBottomCriterionType.topItems 
            : Excel.ConditionalTopBottomCriterionType.bottomItems,
          rank: rank
        };
        break;
        
      case 'custom':
        if (op.criteria) {
          const custom = conditionalFormats.add(Excel.ConditionalFormatType.custom);
          custom.custom.rule.formula = op.criteria;
          
          if (op.format) {
            if (op.format.fontColor) {
              custom.custom.format.font.color = op.format.fontColor;
            }
            if (op.format.fillColor) {
              custom.custom.format.fill.color = op.format.fillColor;
            }
            if (op.format.bold !== undefined) {
              custom.custom.format.font.bold = op.format.bold;
            }
            if (op.format.italic !== undefined) {
              custom.custom.format.font.italic = op.format.italic;
            }
          }
        }
        break;
    }
  }
  
  /**
   * Execute an ADD_COMMENT operation
   */
  private async executeAddComment(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const { sheet, address } = this.parseReference(op.target);
    
    // Use the correct API to add comments - comments are managed at the workbook level
    // Format the reference as SheetName!Address
    const cellReference = `${sheet}!${address}`;
    context.workbook.comments.add(cellReference, op.text);
  }


  /**
   * Execute a SET_FREEZE_PANES operation
   */
  private async executeSetFreezePanes(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    worksheet.activate();
    await context.sync();

    if (op.freeze) {
      // Handle the case where a string address is provided (e.g., "B3")
      if (op.address) {
        // Parse the cell address and freeze at that cell
        const { row, column } = ExcelUtils.parseA1(op.address);
        worksheet.freezePanes.freezeRows(row);
        worksheet.freezePanes.freezeColumns(column);
        await context.sync();
        return;
      }
    }
    else{
      worksheet.freezePanes.unfreeze();
      await context.sync();
      return;
    }
    
    
    
  }

  /**
   * Execute a FORMAT_CHART operation
   */
  private async executeSetWorksheetSettings(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    worksheet.activate();
    await context.sync();
    
    if (op.visible) {
      worksheet.visibility = Excel.SheetVisibility.visible;
    } else {
      worksheet.visibility = Excel.SheetVisibility.hidden;
    }

    if (op.tabColor) {
      worksheet.tabColor = op.tabColor;
    }

    if (op.name) {
      worksheet.name = op.name;
    }

  }

  /**
   * Execute a SET_ACTIVE_SHEET operation
   */
  private async executeSetActiveSheet(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    worksheet.activate();
  }


  /**
   * Execute a FORMAT_CHART operation
   */
  private async executeFormatChart(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    // Get the chart
    let chart: Excel.Chart;
    if (op.chartName || op.chart) {
      // Support both schema's "chart" and interface's "chartName"
      chart = worksheet.charts.getItem(op.chartName || op.chart);
    } else if (op.chartIndex !== undefined) {
      chart = worksheet.charts.getItemAt(op.chartIndex);
    } else {
      throw new Error('Either chartName/chart or chartIndex must be specified');
    }
    
    // Apply data source if specified
    if (op.dataSource) {
      const dataRange = worksheet.getRange(op.dataSource);
      chart.setData(dataRange);
    }
    
    // Chart type
    if (op.hasOwnProperty('chartType') || op.hasOwnProperty('type')) {
      chart.chartType = op.chartType || op.type;
    }
    
    if (op.hasOwnProperty('chartSubType')) {
      // This would need mapping to Office.js API depending on available sub-types
    }
    
    if (op.hasOwnProperty('chartGroup')) {
      // This would need mapping to Office.js API depending on available groups
    }
    
    // Chart position and size
    if (op.hasOwnProperty('height')) {
      chart.height = op.height;
    }
    
    if (op.hasOwnProperty('width')) {
      chart.width = op.width;
    }
    
    if (op.hasOwnProperty('left')) {
      chart.left = op.left;
    }
    
    if (op.hasOwnProperty('top')) {
      chart.top = op.top;
    }
    
    if (op.hasOwnProperty('style')) {
      chart.style = op.style;
    }
    
    // Chart fill
    if (op.hasOwnProperty('fillColor')) {
      chart.format.fill.setSolidColor(op.fillColor);
    }
    
    if (op.hasOwnProperty('hasFill') && !op.hasFill) {
      chart.format.fill.clear();
    }
    
    // Chart border
    if (op.hasOwnProperty('hasBorder')) {
      // @ts-ignore: Property may exist at runtime but not in type definitions
      chart.format.border.visible = op.hasBorder;
    }
    
    if (op.hasOwnProperty('borderColor')) {
      chart.format.border.color = op.borderColor;
    }
    
    if (op.hasOwnProperty('borderWeight')) {
      chart.format.border.weight = op.borderWeight;
    }
    
    if (op.hasOwnProperty('borderStyle')) {
      // Would need mapping to Excel.BorderLineStyle enum
      // Example: chart.format.border.lineStyle = Excel.BorderLineStyle.continuous;
    }
    
    if (op.hasOwnProperty('borderDashStyle')) {
      // Would need mapping to Excel.BorderDashStyle enum
      // Example: chart.format.border.dashStyle = Excel.BorderDashStyle.dash;
    }
    
    // Chart title properties
    if (op.hasOwnProperty('title')) {
      chart.title.text = op.title;
    }
    
    if (op.hasOwnProperty('hasTitle')) {
      chart.title.visible = op.hasTitle;
    }
    
    if (op.hasOwnProperty('titleVisible')) {
      chart.title.visible = op.titleVisible;
    }
    
    if (op.hasOwnProperty('titleColor')) {
      chart.title.format.font.color = op.titleColor;
    }
    
    if (op.hasOwnProperty('titleFontName')) {
      chart.title.format.font.name = op.titleFontName;
    }
    
    if (op.hasOwnProperty('titleFontSize')) {
      chart.title.format.font.size = op.titleFontSize;
    }
    
    if (op.hasOwnProperty('titleFontBold')) {
      chart.title.format.font.bold = op.titleFontBold;
    }
    
    if (op.hasOwnProperty('titleFontItalic')) {
      chart.title.format.font.italic = op.titleFontItalic;
    }
    
    if (op.hasOwnProperty('titlePosition')) {
      // Would need mapping to Office.js position enum if available
    }
    
    // Legend properties
    if (op.hasOwnProperty('hasLegend')) {
      chart.legend.visible = op.hasLegend;
    }
    
    if (op.hasOwnProperty('legendVisible')) {
      chart.legend.visible = op.legendVisible;
    }
    
    if (op.hasOwnProperty('legendPosition')) {
      switch (op.legendPosition) {
        case 'top':
          chart.legend.position = Excel.ChartLegendPosition.top;
          break;
        case 'bottom':
          chart.legend.position = Excel.ChartLegendPosition.bottom;
          break;
        case 'left':
          chart.legend.position = Excel.ChartLegendPosition.left;
          break;
        case 'right':
          chart.legend.position = Excel.ChartLegendPosition.right;
          break;
        case 'corner':
          chart.legend.position = Excel.ChartLegendPosition.corner;
          break;
      }
    }
    
    if (op.hasOwnProperty('legendColor')) {
      chart.legend.format.font.color = op.legendColor;
    }
    
    if (op.hasOwnProperty('legendFontName')) {
      chart.legend.format.font.name = op.legendFontName;
    }
    
    if (op.hasOwnProperty('legendFontSize')) {
      chart.legend.format.font.size = op.legendFontSize;
    }
    
    if (op.hasOwnProperty('legendFontBold')) {
      chart.legend.format.font.bold = op.legendFontBold;
    }
    
    if (op.hasOwnProperty('legendFontItalic')) {
      chart.legend.format.font.italic = op.legendFontItalic;
    }
    
    // Axis properties
    // Get the axis if axis type is specified
    if (op.hasOwnProperty('axisType')) {
      let axis: Excel.ChartAxis;
      
      // Use the getItem method to get the specific axis by type and group
      if (op.axisType === 'category') {
        // For category axis
        axis = chart.axes.getItem(
          Excel.ChartAxisType.category, 
          op.axisGroup === 'secondary' ? Excel.ChartAxisGroup.secondary : Excel.ChartAxisGroup.primary
        );
      } else if (op.axisType === 'value') {
        // For value axis
        axis = chart.axes.getItem(
          Excel.ChartAxisType.value, 
          op.axisGroup === 'secondary' ? Excel.ChartAxisGroup.secondary : Excel.ChartAxisGroup.primary
        );
      } else if (op.axisType === 'series') {
        // For series axis (3D charts)
        axis = chart.axes.getItem(
          Excel.ChartAxisType.series, 
          op.axisGroup === 'secondary' ? Excel.ChartAxisGroup.secondary : Excel.ChartAxisGroup.primary
        );
      }
      
      // Apply axis properties
      if (axis) {
        if (op.hasOwnProperty('title') && op.hasOwnProperty('hasTitle')) {
          axis.title.text = op.title;
          axis.title.visible = op.hasTitle;
        }
        
        if (op.hasOwnProperty('axisVisible')) {
          axis.visible = op.axisVisible;
        }
        
        if (op.hasOwnProperty('axisFontName')) {
          axis.format.font.name = op.axisFontName;
        }
        
        if (op.hasOwnProperty('axisFontSize')) {
          axis.format.font.size = op.axisFontSize;
        }
        
        if (op.hasOwnProperty('axisFontBold')) {
          axis.format.font.bold = op.axisFontBold;
        }
        
        if (op.hasOwnProperty('axisFontItalic')) {
          axis.format.font.italic = op.axisFontItalic;
        }
        
        if (op.hasOwnProperty('axisFontColor')) {
          axis.format.font.color = op.axisFontColor;
        }
        
        if (op.hasOwnProperty('showMajorGridlines')) {
          axis.majorGridlines.visible = op.showMajorGridlines;
        }
        
        if (op.hasOwnProperty('showMinorGridlines')) {
          axis.minorGridlines.visible = op.showMinorGridlines;
        }
        
        if (op.hasOwnProperty('majorUnit')) {
          axis.majorUnit = op.majorUnit;
        }
        
        if (op.hasOwnProperty('minorUnit')) {
          axis.minorUnit = op.minorUnit;
        }
        
        if (op.hasOwnProperty('minimum')) {
          axis.minimum = op.minimum;
        }
        
        if (op.hasOwnProperty('maximum')) {
          axis.maximum = op.maximum;
        }
        
        if (op.hasOwnProperty('displayUnit')) {
          // Would need mapping to Excel.ChartAxisDisplayUnit enum
        }
        
        if (op.hasOwnProperty('logScale')) {
          axis.scaleType = Excel.ChartAxisScaleType.logarithmic;
          axis.logBase = op.logBase ?? 10; // optional numeric base
        }
        
        if (op.hasOwnProperty('reversed')) {
          // Reverse the axis direction
          // @ts-ignore: Property may exist at runtime but not in type definitions
          (axis as any).reversed = op.reversed;
        }
        
        if (op.hasOwnProperty('tickLabelPosition')) {
          // Would need mapping to Excel.ChartAxisTickLabelPosition enum
        }
        
        if (op.hasOwnProperty('tickMarkType')) {
          // Would need mapping to Excel.ChartAxisTickMark enum
        }
      }
    }
    
    // Series properties
    if (op.hasOwnProperty('seriesName') || op.hasOwnProperty('seriesIndex')) {
      let series: Excel.ChartSeries;
      
      if (op.seriesName) {
        // Get series by name if available in Office.js API
        // This might not be directly available - might need to iterate
        const seriesCollection = chart.series;
        // Potential implementation to find by name would go here
      } else if (op.seriesIndex !== undefined) {
        series = chart.series.getItemAt(op.seriesIndex);
      }
      
      if (series) {
        if (op.hasOwnProperty('seriesVisible')) {
          // May not be directly available in Office.js API
        }
        
        if (op.hasOwnProperty('lineColor')) {
          series.format.line.color = op.lineColor;
        }
        
        if (op.hasOwnProperty('lineWeight')) {
          series.format.line.weight = op.lineWeight;
        }
        
        if (op.hasOwnProperty('markerStyle')) {
          // Would need mapping to Excel.ChartMarkerStyle enum
          // series.markerStyle = op.markerStyle;
        }
        
        if (op.hasOwnProperty('markerSize')) {
          series.markerSize = op.markerSize;
        }
        
        if (op.hasOwnProperty('markerColor')) {
          // May need to access through format property
        }
        
        if (op.hasOwnProperty('seriesFillColor')) {
          series.format.fill.setSolidColor(op.seriesFillColor);
        }
        
        if (op.hasOwnProperty('transparency')) {
          // May need conversion to Office.js transparency format
        }
        
        if (op.hasOwnProperty('plotOrder')) {
          // May not be directly available in Office.js API
        }
        
        if (op.hasOwnProperty('gapWidth')) {
          // For certain chart types only
          if (chart.chartType.includes('column') || chart.chartType.includes('bar')) {
            // series.gapWidth = op.gapWidth;
          }
        }
        
        if (op.hasOwnProperty('gapDepth')) {
          // For 3D charts only
        }
      }
    }
    
    // Data point properties
    if (op.hasOwnProperty('seriesIndex') && op.hasOwnProperty('pointIndex')) {
      const series = chart.series.getItemAt(op.seriesIndex);
      if (series && series.points) {
        const point = series.points.getItemAt(op.pointIndex);
        
        if (point) {
          if (op.hasOwnProperty('dataPointVisible')) {
            // May not be directly available in Office.js API
          }
          
          if (op.hasOwnProperty('daatPointfillColor')) { // Note the typo in interface
            point.format.fill.setSolidColor(op.daatPointfillColor);
          }
          
          if (op.hasOwnProperty('dataPointborderColor')) {
            point.format.border.color = op.dataPointborderColor;
          }
          
          if (op.hasOwnProperty('dataPointborderWeight')) {
            point.format.border.weight = op.dataPointborderWeight;
          }
          
          if (op.hasOwnProperty('visible')) {
            try {
              // Attempt to set border visibility - this property may only exist in some Excel versions
              // @ts-ignore: Property may exist at runtime but not in type definitions
              (chart as any).border.visible = op.visible;
            } catch (e) {
              console.warn(`Could not set chart border visibility: ${e}`);
            }
          }
          
          if (op.hasOwnProperty('explosive')) {
            // For pie/doughnut charts
            // point.explosive = op.explosive;
          }
          
          if (op.hasOwnProperty('marker') && op.marker) {
            if (op.marker.style) {
              // Would need mapping to Excel.ChartMarkerStyle enum
            }
            
            if (op.marker.size) {
              point.markerSize = op.marker.size;
            }
            
            if (op.marker.color) {
              // May need to access through format property
            }
          }
        }
      }
    }
    
    // Data label properties
    if (op.hasOwnProperty('dataLabelVisible') || op.hasOwnProperty('dataLabels')) {
      chart.dataLabels.showValue = op.dataLabelVisible || op.dataLabels;
      
      if (op.hasOwnProperty('dataLabelPosition')) {
        // Would need mapping to Excel.ChartDataLabelPosition enum
        // chart.dataLabels.position = Excel.ChartDataLabelPosition[op.dataLabelPosition];
      }
      
      if (op.hasOwnProperty('dataLabelFontName')) {
        chart.dataLabels.format.font.name = op.dataLabelFontName;
      }
      
      if (op.hasOwnProperty('dataLabelFontSize')) {
        chart.dataLabels.format.font.size = op.dataLabelFontSize;
      }
      
      if (op.hasOwnProperty('dataLabelFontBold')) {
        chart.dataLabels.format.font.bold = op.dataLabelFontBold;
      }
      
      if (op.hasOwnProperty('dataLabelFontItalic')) {
        chart.dataLabels.format.font.italic = op.dataLabelFontItalic;
      }
      
      if (op.hasOwnProperty('dataLabelFontColor')) {
        chart.dataLabels.format.font.color = op.dataLabelFontColor;
      }
      
      if (op.hasOwnProperty('dataLabelFormat')) {
        chart.dataLabels.numberFormat = op.dataLabelFormat;
      }
      
      if (op.hasOwnProperty('dataLabelSeparator')) {
        chart.dataLabels.separator = op.dataLabelSeparator;
      }
      
      if (op.hasOwnProperty('dataLabelShowCategoryName')) {
        chart.dataLabels.showCategoryName = op.dataLabelShowCategoryName;
      }
      
      if (op.hasOwnProperty('dataLabelShowSeriesName')) {
        chart.dataLabels.showSeriesName = op.dataLabelShowSeriesName;
      }
      
      if (op.hasOwnProperty('dataLabelShowValue')) {
        chart.dataLabels.showValue = op.dataLabelShowValue;
      }
      
      if (op.hasOwnProperty('dataLabelShowPercentage')) {
        chart.dataLabels.showPercentage = op.dataLabelShowPercentage;
      }
      
      if (op.hasOwnProperty('dataLabelShowBubbleSize')) {
        chart.dataLabels.showBubbleSize = op.dataLabelShowBubbleSize;
      }
    }
  }


  /**
   * Execute a SET_PRINT_SETTINGS operation
   */
  private async executeSetPrintSettings(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    const pageLayout = worksheet.pageLayout;
    
    // Print area settings
    if (op.hasOwnProperty('printAreas') && Array.isArray(op.printAreas) && op.printAreas.length > 0) {
      // Join multiple areas with commas if there are multiple areas
      const printArea = op.printAreas.join(',');
      pageLayout.setPrintArea(printArea);
    }
    
    
    // Print orientation
    if (op.hasOwnProperty('orientation')) {
      pageLayout.orientation = 
        op.orientation === 'landscape' ? Excel.PageOrientation.landscape : Excel.PageOrientation.portrait;
    }
    
    // Print margins
    if (op.hasOwnProperty('topMargin')) pageLayout.topMargin = op.topMargin;
    if (op.hasOwnProperty('rightMargin')) pageLayout.rightMargin = op.rightMargin;
    if (op.hasOwnProperty('bottomMargin')) pageLayout.bottomMargin = op.bottomMargin;
    if (op.hasOwnProperty('leftMargin')) pageLayout.leftMargin = op.leftMargin;
    if (op.hasOwnProperty('headerMargin')) pageLayout.headerMargin = op.headerMargin;
    if (op.hasOwnProperty('footerMargin')) pageLayout.footerMargin = op.footerMargin;
    
    // Print scaling
    if (op.hasOwnProperty('scale') || op.hasOwnProperty('fitToWidth') || op.hasOwnProperty('fitToHeight')) {
      const zoomSettings: any = {};
      
      if (op.hasOwnProperty('scale')) {
        zoomSettings.scale = op.scale;
      }
      
      if (op.hasOwnProperty('fitToWidth')) {
        zoomSettings.horizontalFitToPages = op.fitToWidth;
      }
      
      if (op.hasOwnProperty('fitToHeight')) {
        zoomSettings.verticalFitToPages = op.fitToHeight;
      }
      
      pageLayout.zoom = zoomSettings;
    }
    
    // Print titles
    if (op.hasOwnProperty('printTitles') && Array.isArray(op.printTitles)) {
      // Assume first item is rows, second is columns if both are provided
      if (op.printTitles.length > 0) {
        pageLayout.setPrintTitleRows(op.printTitles[0]);
      }
      
      if (op.printTitles.length > 1) {
        pageLayout.setPrintTitleColumns(op.printTitles[1]);
      }
    }
    
    // Black and white printing
    if (op.hasOwnProperty('blackAndWhite')) {
      pageLayout.blackAndWhite = op.blackAndWhite;
    }
    
    // Draft mode
    if (op.hasOwnProperty('draftMode')) {
      pageLayout.draftMode = op.draftMode;
    }
    
    // Print gridlines
    if (op.hasOwnProperty('printGridlines')) {
      pageLayout.printGridlines = op.printGridlines;
    }
    
    // Print headings
    if (op.hasOwnProperty('headings')) {
      pageLayout.printHeadings = op.headings;
    }
    
    // Print order
    if (op.hasOwnProperty('printOrder')) {
      pageLayout.printOrder = 
        op.printOrder === 'over_then_down' ? Excel.PrintOrder.overThenDown : Excel.PrintOrder.downThenOver;
    }
    
    // Paper size
    if (op.hasOwnProperty('paperSize')) {
      // Map common paper size names to Excel.PaperType enum values
      const paperSizeMap: { [key: string]: Excel.PaperType } = {
        'letter': Excel.PaperType.letter,
        'legal': Excel.PaperType.legal,
        'a3': Excel.PaperType.a3,
        'a4': Excel.PaperType.a4,
        'a5': Excel.PaperType.a5,
        'b4': Excel.PaperType.b4,
        'b5': Excel.PaperType.b5,
        'executive': Excel.PaperType.executive,
        'tabloid': Excel.PaperType.tabloid,
        'statement': Excel.PaperType.statement,
        'envelope10': Excel.PaperType.envelope10,
        'envelopeMonarch': Excel.PaperType.envelopeMonarch,
        'quatro': Excel.PaperType.quatro // Note: correct spelling is 'quatro' not 'quarto'
      };
      
      const paperSize = paperSizeMap[op.paperSize.toLowerCase()] || Excel.PaperType.letter;
      pageLayout.paperSize = paperSize;
    }
    
    // First page number
    if (op.hasOwnProperty('firstPageNumber')) {
      pageLayout.firstPageNumber = op.firstPageNumber;
    }
    
    // Print comments
    if (op.hasOwnProperty('printComments')) {
      // Map printStyle to Excel.PrintComments enum
      const printCommentsMap: { [key: string]: Excel.PrintComments } = {
        'none': Excel.PrintComments.noComments,
        'at_end': Excel.PrintComments.endSheet,
        'as_displayed': Excel.PrintComments.inPlace // Using inPlace instead of displayed
      };
      
      const commentSetting = printCommentsMap[op.printComments] || Excel.PrintComments.noComments;
      pageLayout.printComments = commentSetting;
    }
    
    // Center on page
    if (op.hasOwnProperty('centerHorizontally')) {
      pageLayout.centerHorizontally = op.centerHorizontally;
    }
    
    if (op.hasOwnProperty('centerVertically')) {
      pageLayout.centerVertically = op.centerVertically;
    }
    
    // Headers and footers
    if (op.hasOwnProperty('leftHeader') || op.hasOwnProperty('centerHeader') || op.hasOwnProperty('rightHeader')) {
      try {
        // Access the headers for different page types
        const headersFooters = pageLayout.headersFooters;
        
        if (op.hasOwnProperty('leftHeader')) {
          // Set left header for all page types
          // Access the headers using the appropriate Office.js API
          try {
            // Try different approaches to set headers based on the Office.js version
            if(headersFooters.defaultForAllPages){
              headersFooters.defaultForAllPages.leftHeader = op.leftHeader;
            }
            if (headersFooters.oddPages) headersFooters.oddPages.leftHeader = op.leftHeader;
            if (headersFooters.firstPage) headersFooters.firstPage.leftHeader = op.leftHeader;
            if (headersFooters.evenPages) headersFooters.evenPages.leftHeader = op.leftHeader;
          } catch (error) {
            console.warn('Error setting left header:', error);
          }
        }
        
        if (op.hasOwnProperty('centerHeader')) {
          // Set center header for all page types
          try {
            if(headersFooters.defaultForAllPages){
              headersFooters.defaultForAllPages.centerHeader = op.centerHeader;
            }
            if (headersFooters.oddPages) headersFooters.oddPages.centerHeader = op.centerHeader;
            if (headersFooters.firstPage) headersFooters.firstPage.centerHeader = op.centerHeader;
            if (headersFooters.evenPages) headersFooters.evenPages.centerHeader = op.centerHeader;
          } catch (error) {
            console.warn('Error setting center header:', error);
          }
        }
        
        if (op.hasOwnProperty('rightHeader')) {
          // Set right header for all page types
          try {
            if(headersFooters.defaultForAllPages){
              headersFooters.defaultForAllPages.rightHeader = op.rightHeader;
            }
            if (headersFooters.oddPages) headersFooters.oddPages.rightHeader = op.rightHeader;
            if (headersFooters.firstPage) headersFooters.firstPage.rightHeader = op.rightHeader;
            if (headersFooters.evenPages) headersFooters.evenPages.rightHeader = op.rightHeader;
          } catch (error) {
            console.warn('Error setting right header:', error);
          }
        }
      } catch (error) {
        console.error('Error setting headers:', error);
      }
    }
    
    if (op.hasOwnProperty('leftFooter') || op.hasOwnProperty('centerFooter') || op.hasOwnProperty('rightFooter')) {
      try {
        // Access the footers for different page types
        const headersFooters = pageLayout.headersFooters;
        
        if (op.hasOwnProperty('leftFooter')) {
          // Set left footer for all page types
          try {
            if(headersFooters.defaultForAllPages){
              headersFooters.defaultForAllPages.leftFooter = op.leftFooter;
            }
            if (headersFooters.oddPages) headersFooters.oddPages.leftFooter = op.leftFooter;
            if (headersFooters.firstPage) headersFooters.firstPage.leftFooter = op.leftFooter;
            if (headersFooters.evenPages) headersFooters.evenPages.leftFooter = op.leftFooter;
          } catch (error) {
            console.warn('Error setting left footer:', error);
          }
        }
        
        if (op.hasOwnProperty('centerFooter')) {
          // Set center footer for all page types
          try {
            if(headersFooters.defaultForAllPages){
              headersFooters.defaultForAllPages.centerFooter = op.centerFooter;
            }
            if (headersFooters.oddPages) headersFooters.oddPages.centerFooter = op.centerFooter;
            if (headersFooters.firstPage) headersFooters.firstPage.centerFooter = op.centerFooter;
            if (headersFooters.evenPages) headersFooters.evenPages.centerFooter = op.centerFooter;
          } catch (error) {
            console.warn('Error setting center footer:', error);
          }
        }
        
        if (op.hasOwnProperty('rightFooter')) {
          // Set right footer for all page types
          try {
            if(headersFooters.defaultForAllPages){
              headersFooters.defaultForAllPages.rightFooter = op.rightFooter;
            }
            if (headersFooters.oddPages) headersFooters.oddPages.rightFooter = op.rightFooter;
            if (headersFooters.firstPage) headersFooters.firstPage.rightFooter = op.rightFooter;
            if (headersFooters.evenPages) headersFooters.evenPages.rightFooter = op.rightFooter;
          } catch (error) {
            console.warn('Error setting right footer:', error);
          }
        }
      } catch (error) {
        console.error('Error setting footers:', error);
      }
    }
  }

  /**
   * Execute a SET_PAGE_SETUP operation
   */
  private async executeSetPageSetup(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    try {
      // Set zoom level
      if (op.hasOwnProperty('zoom')) {
        // Set zoom using pageLayout.zoom with a scale property
        worksheet.pageLayout.zoom = { scale: op.zoom };

        await context.sync();
        
        // Note: Setting zoom may require different approach depending on Office.js version
        console.log(`Setting zoom level to ${op.zoom}%`);
      }
      
      // Set gridlines visibility
      if (op.hasOwnProperty('gridlines')) {
        worksheet.showGridlines = op.gridlines;
      }
      
      // Set headers visibility
      if (op.hasOwnProperty('headers')) {
        worksheet.showHeadings = op.headers;
      }
      
      // Set page layout view
      if (op.hasOwnProperty('pageLayoutView')) {
        if (op.pageLayoutView === 'print') {
          worksheet.activate();
          // Use the correct method to show print preview
          // Note: This may vary depending on Office.js version
          try {
            // Try to use the Office UI namespace if available
            Office.context.ui.displayDialogAsync('https://outlook.office.com/mail/printview', 
              { height: 80, width: 80, displayInIframe: true });
          } catch (error) {
            console.error('Unable to display print preview:', error);
          }
        }
      }
    } catch (error) {
      console.error('Error setting page setup:', error);
    }
  }

  /**
   * Execute a FORMAT_CHART_SERIES operation
   */
  private async executeFormatChartSeries(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    // Get the chart
    let chart: Excel.Chart;
    if (op.chartName) {
      chart = worksheet.charts.getItem(op.chartName);
    } else if (op.chartIndex !== undefined) {
      chart = worksheet.charts.getItemAt(op.chartIndex);
    } else {
      throw new Error('Either chartName or chartIndex must be specified');
    }
    
    // Get the series
    let series: Excel.ChartSeries;
    if (op.seriesName) {
      series = chart.series.getItemAt(0); // Placeholder - the API doesn't directly support getting by name
      // We'd need to loop through series to find by name
    } else if (op.seriesIndex !== undefined) {
      series = chart.series.getItemAt(op.seriesIndex);
    } else {
      throw new Error('Either seriesName or seriesIndex must be specified');
    }
    
    // Apply formatting
    if (op.hasOwnProperty('lineColor')) {
      series.format.line.color = op.lineColor;
    }
    
    if (op.hasOwnProperty('lineWeight')) {
      series.format.line.weight = op.lineWeight;
    }
    
    if (op.hasOwnProperty('markerStyle')) {
      series.markerStyle = op.markerStyle;
    }
    
    if (op.hasOwnProperty('markerSize')) {
      series.markerSize = op.markerSize;
    }
    
    if (op.hasOwnProperty('fillColor')) {
      series.format.fill.setSolidColor(op.fillColor);
    }
  }

  /**
   * Execute a COMPOSITE_OPERATION operation
   */
  private async executeCompositeOperation(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const compositeOp = operation as any;
    console.log(`Executing composite operation: ${compositeOp.name}`);
    
    try {
      for (const subOp of compositeOp.subOperations) {
        try {
          await this.executeOperation(context, subOp);
        } catch (error) {
          console.error(`Error executing sub-operation: ${error.message}`);
          if (compositeOp.abortOnFailure) {
            throw error;
          }
        }
      }
    } catch (error) {
      console.error(`Error executing composite operation: ${error.message}`);
      throw error;
    }
  }

  /**
   * Execute a BATCH_OPERATION operation
   */
  private async executeBatchOperation(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const batchOp = operation as any;
    console.log(`Executing batch operation with ${batchOp.operations.length} operations`);
    
    for (const op of batchOp.operations) {
      await this.executeOperation(context, op);
    }
    
    if (batchOp.requiresSync) {
      await context.sync();
    }
  }
  
  /**
   * Execute an EXPORT_TO_PDF operation
   */
  private async executeExportToPdf(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    try {
      // Activate the worksheet to ensure it's the active one for export
      worksheet.activate();
      await context.sync();      
      
      
      // Export to PDF
      // Note: The exact implementation depends on the Office.js version and platform
      // This is a general approach that may need adaptation
      
      console.log(`Preparing to export worksheet ${op.sheet} to PDF`);
      
      // Use the Office.js API to export to PDF
      // Since the exact API varies by Office.js version, we'll implement a few approaches
      
      // First, check if we're in a context where we can use the Office UI to save as PDF
      try {
        // Approach 1: Use the Office UI dialog to save as PDF
        // This is the most widely supported approach across different Office versions
        console.log('Attempting to export using Office UI...');
        
        // Set up the worksheet for printing
        await context.sync();
        
        // Use the Office UI to show the save dialog
        // Note: This will prompt the user to save the file
        Office.context.ui.displayDialogAsync(
          'https://appsforoffice.microsoft.com/lib/1.1/hosted/office-js/pdf-dialog.html',
          { height: 50, width: 50, displayInIframe: true },
        );
      } catch (error) {
        console.error('Error exporting to PDF:', error);
      }
    }
    catch (error) {
      console.error('Error exporting to PDF:', error);
    }
  }

  /**
   * Execute a CREATE_SCENARIO_TABLE operation
   */

  private async executeCreateScenarioTable(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as CreateScenarioTableOperation;
    
    // Get the target range for the scenario table
    const { sheet: sheetName, address: rangeAddress } = this.parseReference(op.range);
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const tableRange = worksheet.getRange(rangeAddress);
    
    // Get references to formula and input cells
    const { sheet: formulaSheet, address: formulaAddress } = this.parseReference(op.formulaCell);
    const formulaCell = context.workbook.worksheets.getItem(formulaSheet).getRange(formulaAddress);
    
    const { sheet: inputSheet, address: inputAddress } = this.parseReference(op.inputCell);
    const inputCell = context.workbook.worksheets.getItem(inputSheet).getRange(inputAddress);
    
    // Load current values
    formulaCell.load("formulas, values");
    inputCell.load("values");
    await context.sync();
    
    const originalInputValue = inputCell.values[0][0];
    const formulaValue = formulaCell.values[0][0];
    const hasFormula = typeof formulaCell.formulas[0][0] === "string" && 
                       formulaCell.formulas[0][0].toString().startsWith("=");
    
    // Get the formula text (for potential display)
    const formulaText = hasFormula ? formulaCell.formulas[0][0].toString() : "";
    
    // Determine if it's a one-way or two-way table
    const isOneWay = !op.tableType || op.tableType === 'one-way';
    
    let rowInputSheet = "";
    let rowInputAddress = "";
    let originalRowInputValue: any = null;
    let rowInputCell: Excel.Range | null = null;
    
    // For two-way tables, get the row input cell information
    if (!isOneWay) {
      // Set up two-way table
      if (!op.rowInputCell || !op.rowValues || op.rowValues.length === 0) {
        throw new Error("Two-way scenario table requires rowInputCell and rowValues");
      }
      
      const rowCellInfo = this.parseReference(op.rowInputCell);
      rowInputSheet = rowCellInfo.sheet;
      rowInputAddress = rowCellInfo.address;
      rowInputCell = context.workbook.worksheets.getItem(rowInputSheet).getRange(rowInputAddress);
      
      rowInputCell.load("values");
      await context.sync();
      
      originalRowInputValue = rowInputCell.values[0][0];
    }
    
    if (isOneWay) {
      // Set up one-way table
      await this.createOneWayScenarioTable(
        context, 
        tableRange, 
        formulaCell, 
        inputCell, 
        op.values, 
        formulaText, 
        op.format,
        op.includeFormula || false
      );
    } else if (rowInputCell) {
      await this.createTwoWayScenarioTable(
        context,
        tableRange,
        formulaCell,
        inputCell,
        rowInputCell,
        op.values,
        op.rowValues!,
        op.format,
        op.includeFormula || false,
        formulaText
      );
    }
    
    // Run scenarios if requested
    if (op.runScenarios) {
      await this.runScenariosForTable(
        context,
        isOneWay,
        inputCell,
        isOneWay ? null : rowInputCell,
        op.values,
        isOneWay ? null : op.rowValues
      );
    }
    
    // Restore original input values
    inputCell.values = [[originalInputValue]];
    if (!isOneWay && rowInputCell) {
      rowInputCell.values = [[originalRowInputValue]];
    }
    
    await context.sync();
  }
  
  /**
   * Create a one-way scenario table
   */
  private async createOneWayScenarioTable(
    context: Excel.RequestContext,
    tableRange: Excel.Range,
    formulaCell: Excel.Range,
    inputCell: Excel.Range,
    values: (string | number)[],
    formulaText: string,
    format?: { headerFormatting?: boolean, resultFormatting?: string, tableStyle?: string },
    includeFormula: boolean = false
  ): Promise<void> {
    // Load table dimensions
    tableRange.load("rowCount, columnCount");
    await context.sync();
    
    // Check if the range is big enough
    const requiredRows = values.length + 1; // +1 for header
    const requiredColumns = includeFormula ? 3 : 2; // Input value, result, (formula)
    
    if (tableRange.rowCount < requiredRows || tableRange.columnCount < requiredColumns) {
      throw new Error(`Table range too small. Needs at least ${requiredRows} rows and ${requiredColumns} columns.`);
    }
    
    // Set up headers
    const headers = tableRange.getRow(0);
    // Set header values with proper array structure
    if (includeFormula) {
      headers.values = [["Input Value", "Result", "Formula"]];
    } else {
      headers.values = [["Input Value", "Result"]];
    }
    
    if (format?.headerFormatting) {
      headers.format.font.bold = true;
    }
    
    // Get formula text and references for the sticky-IF formulas
    const formulaCellAddress = formulaCell.address;
    const inputCellAddress = inputCell.address;
    
    // Build each row of the table
    for (let i = 0; i < values.length; i++) {
      const row = tableRange.getRow(i + 1);
      const testValue = values[i];
      
      // Set input value in first column
      const inputValueCell = row.getCell(0, 0);
      inputValueCell.values = [[testValue]];
      
      // Create sticky-IF formula in second column
      const stickyFormula = 
        `=IF(${inputCellAddress}=${testValue}, ${formulaCellAddress}, ` +
        `IF(ISBLANK(INDIRECT("RC",FALSE)), ${formulaCellAddress}, INDIRECT("RC",FALSE)))`;
      
      const resultCell = row.getCell(0, 1);
      resultCell.formulas = [[stickyFormula]];
      
      // Add formula text if requested
      if (includeFormula && formulaText) {
        const formulaCell = row.getCell(0, 2);
        // Ensure we're using array structure for cell values
        formulaCell.values = [[formulaText.toString()]];
      }
    }
    
    // Apply table style if specified
    if (format?.tableStyle) {
      tableRange.format.autofitColumns();
      // Apply a table style - implementation depends on your requirements
      // Could use predefined table styles or custom formatting
    }
    
    await context.sync();
  }
  
  /**
   * Create a two-way scenario table
   */
  private async createTwoWayScenarioTable(
    context: Excel.RequestContext,
    tableRange: Excel.Range,
    formulaCell: Excel.Range,
    columnInputCell: Excel.Range,
    rowInputCell: Excel.Range,
    columnValues: (string | number)[],
    rowValues: (string | number)[],
    format?: { headerFormatting?: boolean, resultFormatting?: string, tableStyle?: string },
    includeFormula: boolean = false,
    formulaText: string = ""
  ): Promise<void> {
    // Load table dimensions
    tableRange.load("rowCount, columnCount");
    await context.sync();
    
    // Check if the range is big enough
    const requiredRows = rowValues.length + 1; // +1 for header
    // If including formula text, need an extra column for the original formula
    const requiredColumns = includeFormula ? columnValues.length + 2 : columnValues.length + 1;
    
    if (tableRange.rowCount < requiredRows || tableRange.columnCount < requiredColumns) {
      throw new Error(`Table range too small. Needs at least ${requiredRows} rows and ${requiredColumns} columns.`);
    }
    
    // Set up top-left corner cell
    const topLeftCell = tableRange.getCell(0, 0);
    topLeftCell.values = [[""]]; // Empty string in the top-left cell
    
    // Set up column headers (column input values)
    for (let c = 0; c < columnValues.length; c++) {
      const headerValue = columnValues[c];
      tableRange.getCell(0, c + 1).values = [[headerValue]];
    }
    
    // Set up row headers (row input values)
    for (let r = 0; r < rowValues.length; r++) {
      const headerValue = rowValues[r];
      tableRange.getCell(r + 1, 0).values = [[headerValue]];
    }
    
    // Apply header formatting if specified
    if (format?.headerFormatting) {
      tableRange.getRow(0).format.font.bold = true;
      tableRange.getColumn(0).format.font.bold = true;
    }
    
    // Get references for the sticky-IF formulas
    const formulaCellAddress = formulaCell.address;
    const columnInputCellAddress = columnInputCell.address;
    const rowInputCellAddress = rowInputCell.address;
    
    // Build each cell of the table with sticky-IF formulas
    for (let r = 0; r < rowValues.length; r++) {
      for (let c = 0; c < columnValues.length; c++) {
        const rowValue = rowValues[r];
        const colValue = columnValues[c];
        
        // Create complex sticky-IF formula for two-way table
        const stickyFormula = 
          `=IF(AND(${columnInputCellAddress}=${colValue}, ${rowInputCellAddress}=${rowValue}), ${formulaCellAddress}, ` +
          `IF(ISBLANK(INDIRECT("RC",FALSE)), ${formulaCellAddress}, INDIRECT("RC",FALSE)))`;
        
        const resultCell = tableRange.getCell(r + 1, c + 1);
        resultCell.formulas = [[stickyFormula]];
        
      }
    }
    
    // Apply table style if specified
    if (format?.tableStyle) {
      tableRange.format.autofitColumns();
      // Apply a table style - implementation depends on your requirements
    }
    
    // Add formula text if requested
    if (includeFormula && formulaText) {
      try {
        // Add formula in a separate cell (e.g., below the table or in a specific location)
        const formulaInfoCell = tableRange.getCell(rowValues.length + 1, 0);
        formulaInfoCell.values = [["Formula:"]];
        formulaInfoCell.format.font.bold = true;
        
        const formulaDisplayCell = tableRange.getCell(rowValues.length + 1, 1);
        formulaDisplayCell.values = [[formulaText.toString()]];
      } catch (e) {
        console.warn(`Could not add formula text: ${e}`);
      }
    }
    
    await context.sync();
  }
  
  /**
   * Run through all scenarios to populate the table
   */
  private async runScenariosForTable(
    context: Excel.RequestContext,
    isOneWay: boolean,
    inputCell: Excel.Range,
    rowInputCell: Excel.Range | null,
    columnValues: (string | number)[],
    rowValues: (string | number)[] | null
  ): Promise<void> {
    if (isOneWay) {
      // One-way table: cycle through column values
      for (const value of columnValues) {
        inputCell.values = [[value]];
        await context.sync();
        
        // Pause briefly to allow calculation
        await new Promise(resolve => setTimeout(resolve, 150));
      }
    } else {
      // Two-way table: cycle through all combinations
      for (const rowValue of rowValues!) {
        rowInputCell!.values = [[rowValue]];
        
        for (const colValue of columnValues) {
          inputCell.values = [[colValue]];
          await context.sync();
          
          // Pause briefly to allow calculation
          await new Promise(resolve => setTimeout(resolve, 150));
        }
      }
    }
  }

    /**
   * Execute a SET_ROW_COLUMN_OPTIONS operation
   * Sets row or column properties like size, grouping and visibility
   */
  private async executeSetRowColumnOptions(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    try {
      const op = operation as SetRowColumnOptionsOperation;
      const worksheet = context.workbook.worksheets.getItem(op.sheet);
      
      // Function to process a range of rows or columns
      const processItems = async (type: 'row' | 'column', indices: number[]) => {
        // Handle size options first
        if (op.size !== undefined || op.autofit) {
          for (const index of indices) {
            if (type === 'row') {
              // Get the entire row as a range
              const rowRange = worksheet.getRanges(`${index+1}:${index+1}`);
              
              if (op.size !== undefined) {
                rowRange.format.rowHeight = op.size;
              }
              
              if (op.autofit) {
                rowRange.format.autofitRows();
              }
            } else { // column
              // Convert zero-based index to column letter (A, B, C, etc.)
              const colLetter = this.columnIndexToLetter(index);
              const colRange = worksheet.getRanges(`${colLetter}:${colLetter}`);
              
              if (op.size !== undefined) {
                colRange.format.columnWidth = op.size;
              }
              
              if (op.autofit) {
                colRange.format.autofitColumns();
              }
            }
          }
        }
        
        // Handle visibility options
        if (op.hidden !== undefined) {
          for (const index of indices) {
            if (type === 'row') {
              const rowRange = worksheet.getRange(`${index+1}:${index+1}`);
              rowRange.rowHidden = op.hidden;
            } else { // column
              const colLetter = this.columnIndexToLetter(index);
              const colRange = worksheet.getRange(`${colLetter}:${colLetter}`);
              colRange.columnHidden = op.hidden;
            }
          }
        }
        
        // Handle group operations
        if (op.group && op.group.length > 0) {
          // Excel API requires syncing before grouping operations
          await context.sync();
          
          for (const group of op.group) {
            // Validate group indices
            if (group.start < 0 || group.end < group.start) {
              console.warn(`Invalid group range: ${group.start}-${group.end}`);
              continue;
            }
            
            try {
              if (type === 'row') {
                // Get row range for grouping
                const startRow = group.start + 1; // 1-based
                const endRow = group.end + 1;     // 1-based
                const groupRange = worksheet.getRange(`${startRow}:${endRow}`);
                
                // Apply grouping
                groupRange.group("ByRows");

              } else { // column
                // Get column range for grouping
                const startCol = this.columnIndexToLetter(group.start);
                const endCol = this.columnIndexToLetter(group.end);
                const groupRange = worksheet.getRange(`${startCol}:${endCol}`);
                
                // Apply grouping 
                groupRange.group("ByColumns");

              }
            } catch (e) {
              console.error(`Error grouping ${type}s ${group.start}-${group.end}: ${e}`);
            }
          }
        }
      };
      
      // Process the rows or columns based on type
      await processItems(op.type, op.indices);
      
    } catch (error) {
      console.error(`Error setting row/column options: ${error}`);
      throw error;
    }
  }

  // Helper method to convert column index to letter
  private columnIndexToLetter(index: number): string {
    let columnName = '';
    let dividend = index + 1;
    let modulo: number;
    
    while (dividend > 0) {
      modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }
    
    return columnName;
  }

  private async executeSetCalculationOptions(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    try {
      const op = operation as SetCalculationOptionsOperation;
      
      // Set calculation mode if specified
      if (op.calculationMode !== undefined) {
        context.workbook.application.calculationMode = op.calculationMode;
      }
      
      // Set iterative calculation options
      if (op.iterative !== undefined) {
        context.workbook.application.iterativeCalculation.enabled = op.iterative;
      }
      
      if (op.maxIterations !== undefined) {
        context.workbook.application.iterativeCalculation.maxIteration = op.maxIterations;
      }
      
      if (op.maxChange !== undefined) {
        context.workbook.application.iterativeCalculation.maxChange = op.maxChange;
      }
      
      // Force calculation if requested
      if (op.calculate) {
        const calcType = op.calculationType || Excel.CalculationType.full;
        context.workbook.application.calculate(calcType);
      }
      
    } catch (error) {
      console.error(`Error setting calculation options: ${error}`);
      throw error;
    }
  }

    /**
   * Execute a RECALCULATE operation to refresh calculations
   * Can target the entire workbook, specific worksheets, or specific ranges
   */
  private async executeRecalculateRanges(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    try {
      const op = operation as RecalculateRangesOperation;
      
      // Case 1: Recalculate entire workbook
      if (op.recalculateAll) {
        // Use Full for complete recalculation including dependencies
        context.workbook.application.calculate(Excel.CalculationType.full);
        return;
      }
      
      // Case 2: Recalculate specific worksheets
      if (op.sheets && op.sheets.length > 0) {
        for (const sheetName of op.sheets) {
          try {
            const worksheet = context.workbook.worksheets.getItem(sheetName);
            worksheet.calculate(true);
          } catch (err) {
            console.error(`Error recalculating worksheet ${sheetName}:`, err);
          }
        }
      }
      
      // Case 3: Recalculate specific ranges
      if (op.ranges && op.ranges.length > 0) {
        for (const rangeRef of op.ranges) {
          try {
            // Parse the reference into sheet and address
            const { sheet, address } = this.parseReference(rangeRef);
            const worksheet = context.workbook.worksheets.getItem(sheet);
            const range = worksheet.getRange(address);
            
            // Calculate just this range
            range.calculate();
          } catch (err) {
            console.error(`Error recalculating range ${rangeRef}:`, err);
          }
        }
      }
      
      // Case 4: Calculate a specific cell
      if (op.cell) {
        try {
          const { sheet, address } = this.parseReference(op.cell);
          const worksheet = context.workbook.worksheets.getItem(sheet);
          const cell = worksheet.getRange(address);
          cell.calculate();
        } catch (err) {
          console.error(`Error recalculating cell ${op.cell}:`, err);
        }
      }
      
      // If nothing specific was provided, calculate everything
      if ((!op.sheets || op.sheets.length === 0) && 
          (!op.ranges || op.ranges.length === 0) && 
          !op.cell && 
          !op.recalculateAll) {
        context.workbook.application.calculate(Excel.CalculationType.full);
      }
      
    } catch (error) {
      console.error(`Error during recalculation:`, error);
      throw error;
    }
  }
}


