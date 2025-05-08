/**
 * Undo Handlers
 * 
 * Provides functionality to undo Excel operations by restoring the previous state.
 * Each operation type has a specific undo handler tailored to its behavior.
 */

import { ExcelOperation, ExcelOperationType } from '../../models/ExcelOperationModels';
import { WorkbookAction, AffectedRange, BeforeState } from '../../models/VersionModels';

/**
 * Service for handling undo operations for different Excel operation types
 */
export class UndoHandlers {
  /**
   * Undo an action by restoring its before state
   * @param context The Excel context
   * @param action The action to undo
   * @returns A promise that resolves when the undo is complete
   */
  public async undoAction(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    try {
      console.log(`🔄 [UndoHandlers] Starting to undo action: ${action.id}`);
      console.log(`📋 [UndoHandlers] Action details:`, {
        id: action.id,
        type: action.operation?.op,
        description: action.description,
        affectedRanges: action.affectedRanges?.map(r => `${r.sheetName}!${r.range}`),
        hasBeforeState: !!action.beforeState
      });
      
      // Choose the appropriate undo handler based on operation type
      const operationType = action.operation.op as ExcelOperationType;
      console.log(`🔄 [UndoHandlers] Using handler for operation type: ${operationType}`);
      
      switch (operationType) {
        case ExcelOperationType.SET_VALUE:
          await this.undoSetValue(context, action);
          break;
          
        case ExcelOperationType.ADD_FORMULA:
          await this.undoAddFormula(context, action);
          break;
          
        case ExcelOperationType.FORMAT_RANGE:
          await this.undoFormatRange(context, action);
          break;
          
        case ExcelOperationType.CLEAR_RANGE:
          await this.undoClearRange(context, action);
          break;
          
        case ExcelOperationType.CREATE_TABLE:
          await this.undoCreateTable(context, action);
          break;
          
        case ExcelOperationType.SORT_RANGE:
          await this.undoSortRange(context, action);
          break;
          
        case ExcelOperationType.FILTER_RANGE:
          await this.undoFilterRange(context, action);
          break;
          
        case ExcelOperationType.CREATE_SHEET:
          await this.undoCreateSheet(context, action);
          break;
          
        case ExcelOperationType.DELETE_SHEET:
          await this.undoDeleteSheet(context, action);
          break;
          
        case ExcelOperationType.RENAME_SHEET:
          await this.undoRenameSheet(context, action);
          break;
          
        case ExcelOperationType.COPY_RANGE:
          await this.undoCopyRange(context, action);
          break;
          
        case ExcelOperationType.MERGE_CELLS:
          await this.undoMergeCells(context, action);
          break;
          
        case ExcelOperationType.UNMERGE_CELLS:
          await this.undoUnmergeCells(context, action);
          break;
          
        case ExcelOperationType.CREATE_CHART:
          await this.undoCreateChart(context, action);
          break;
          
        case ExcelOperationType.COMPOSITE_OPERATION:
        case ExcelOperationType.BATCH_OPERATION:
          await this.undoCompositeOperation(context, action);
          break;
          
        default:
          console.warn(`No specific undo handler for operation type: ${action.operation.op}`);
          await this.undoGeneric(context, action);
      }
    } catch (error) {
      console.error(`Error undoing action ${action.id}:`, error);
      throw error;
    }
  }
  
  /**
   * Undo a setValue operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoSetValue(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.values) {
      console.warn('Cannot undo setValue: missing affected ranges or before state');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Restore the previous value
    targetRange.values = action.beforeState.values;
    
    await context.sync();
  }
  
  /**
   * Undo an addFormula operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoAddFormula(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.formulas) {
      console.warn('Cannot undo addFormula: missing affected ranges or before state');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Restore the previous formula
    targetRange.formulas = action.beforeState.formulas;
    
    await context.sync();
  }
  
  /**
   * Undo a formatRange operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoFormatRange(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.formats || !action.beforeState.formats.length) {
      console.warn('Cannot undo formatRange: missing affected ranges or before state');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Get the format from before state
    const format = action.beforeState.formats[0];
    
    // Restore the previous formatting
    if (format.fontFamily) targetRange.format.font.name = format.fontFamily;
    if (format.fontSize) targetRange.format.font.size = format.fontSize;
    if (format.bold !== undefined) targetRange.format.font.bold = format.bold;
    if (format.italic !== undefined) targetRange.format.font.italic = format.italic;
    if (format.underline !== undefined) targetRange.format.font.underline = format.underline;
    if (format.fontColor) targetRange.format.font.color = format.fontColor;
    if (format.fillColor) targetRange.format.fill.color = format.fillColor;
    if (format.horizontalAlignment) targetRange.format.horizontalAlignment = format.horizontalAlignment;
    if (format.verticalAlignment) targetRange.format.verticalAlignment = format.verticalAlignment;
    
    await context.sync();
  }
  
  /**
   * Undo a clearRange operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoClearRange(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.values) {
      console.warn('Cannot undo clearRange: missing affected ranges or before state');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Restore values and formulas
    if (action.beforeState.values) targetRange.values = action.beforeState.values;
    if (action.beforeState.formulas) targetRange.formulas = action.beforeState.formulas;
    
    // Restore formatting if available
    if (action.beforeState.formats && action.beforeState.formats.length > 0) {
      const format = action.beforeState.formats[0];
      
      if (format.fontFamily) targetRange.format.font.name = format.fontFamily;
      if (format.fontSize) targetRange.format.font.size = format.fontSize;
      if (format.bold !== undefined) targetRange.format.font.bold = format.bold;
      if (format.italic !== undefined) targetRange.format.font.italic = format.italic;
      if (format.fontColor) targetRange.format.font.color = format.fontColor;
      if (format.fillColor) targetRange.format.fill.color = format.fillColor;
    }
    
    await context.sync();
  }
  
  /**
   * Undo a createTable operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoCreateTable(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo createTable: missing affected ranges');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    
    // Find tables in the affected range
    const tables = worksheet.tables;
    tables.load('items');
    await context.sync();
    
    // Delete any tables in the range
    for (const table of tables.items) {
      table.delete();
    }
    
    // Restore the previous values and formatting
    if (action.beforeState.values || action.beforeState.formulas) {
      const targetRange = worksheet.getRange(range);
      
      if (action.beforeState.values) targetRange.values = action.beforeState.values;
      if (action.beforeState.formulas) targetRange.formulas = action.beforeState.formulas;
      
      // Restore formatting if available
      if (action.beforeState.formats && action.beforeState.formats.length > 0) {
        const format = action.beforeState.formats[0];
        
        if (format.fontFamily) targetRange.format.font.name = format.fontFamily;
        if (format.fontSize) targetRange.format.font.size = format.fontSize;
        if (format.bold !== undefined) targetRange.format.font.bold = format.bold;
        if (format.italic !== undefined) targetRange.format.font.italic = format.italic;
        if (format.fontColor) targetRange.format.font.color = format.fontColor;
        if (format.fillColor) targetRange.format.fill.color = format.fillColor;
      }
    }
    
    await context.sync();
  }
  
  /**
   * Undo a sortRange operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoSortRange(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.values) {
      console.warn('Cannot undo sortRange: missing affected ranges or before state');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Simply restore the previous values and formulas
    if (action.beforeState.values) targetRange.values = action.beforeState.values;
    if (action.beforeState.formulas) targetRange.formulas = action.beforeState.formulas;
    
    await context.sync();
  }
  
  /**
   * Undo a filterRange operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoFilterRange(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo filterRange: missing affected ranges');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    
    // Find tables in the affected range
    const tables = worksheet.tables;
    tables.load('items');
    await context.sync();
    
    // Clear filters on any tables in the range
    for (const table of tables.items) {
      table.clearFilters();
    }
    
    await context.sync();
  }
  
  /**
   * Undo a createSheet operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoCreateSheet(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo createSheet: missing affected ranges');
      return;
    }
    
    const { sheetName } = action.affectedRanges[0];
    
    // Delete the created sheet
    try {
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      worksheet.delete();
      await context.sync();
    } catch (error) {
      console.warn(`Error deleting sheet ${sheetName}:`, error);
    }
  }
  
  /**
   * Undo a deleteSheet operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoDeleteSheet(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.sheetProperties) {
      console.warn('Cannot undo deleteSheet: missing affected ranges or before state');
      return;
    }
    
    const { sheetName } = action.affectedRanges[0];
    const sheetProps = action.beforeState.sheetProperties[sheetName];
    
    if (!sheetProps) {
      console.warn(`Cannot undo deleteSheet: missing properties for sheet ${sheetName}`);
      return;
    }
    
    // Recreate the deleted sheet
    try {
      // Add the sheet at the original position if possible
      const newSheet = context.workbook.worksheets.add(sheetName);
      
      if (sheetProps.position !== undefined) {
        newSheet.position = sheetProps.position;
      }
      
      // Set visibility if specified
      if (sheetProps.visibility !== undefined) {
        newSheet.visibility = sheetProps.visibility;
      }
      
      await context.sync();
      
      // Note: We cannot restore the sheet contents as they weren't captured
      // A full sheet capture would be too expensive for normal operations
      console.log(`Recreated sheet ${sheetName} but could not restore its contents`);
    } catch (error) {
      console.error(`Error recreating sheet ${sheetName}:`, error);
    }
  }
  
  /**
   * Undo a renameSheet operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoRenameSheet(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length || !action.beforeState.sheetProperties) {
      console.warn('Cannot undo renameSheet: missing affected ranges or before state');
      return;
    }
    
    const { sheetName } = action.affectedRanges[0];
    
    // Find the original name from the operation
    const operation = action.operation;
    if (!('name' in operation)) {
      console.warn('Cannot undo renameSheet: missing name in operation');
      return;
    }
    
    // Get the original name from before state if available
    let originalName = '';
    for (const [name, props] of Object.entries(action.beforeState.sheetProperties)) {
      if (props && 'name' in props) {
        originalName = name;
        break;
      }
    }
    
    if (!originalName) {
      console.warn('Cannot undo renameSheet: could not determine original name');
      return;
    }
    
    // Rename the sheet back to its original name
    try {
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      worksheet.name = originalName;
      await context.sync();
    } catch (error) {
      console.error(`Error renaming sheet ${sheetName} back to ${originalName}:`, error);
    }
  }
  
  /**
   * Undo a copyRange operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoCopyRange(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (action.affectedRanges.length < 2 || !action.beforeState.values) {
      console.warn('Cannot undo copyRange: missing affected ranges or before state');
      return;
    }
    
    // The destination range is the second affected range
    const { sheetName, range } = action.affectedRanges[1];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Restore the previous values at the destination
    if (action.beforeState.values) targetRange.values = action.beforeState.values;
    if (action.beforeState.formulas) targetRange.formulas = action.beforeState.formulas;
    
    // Restore formatting if available
    if (action.beforeState.formats && action.beforeState.formats.length > 0) {
      const format = action.beforeState.formats[0];
      
      if (format.fontFamily) targetRange.format.font.name = format.fontFamily;
      if (format.fontSize) targetRange.format.font.size = format.fontSize;
      if (format.bold !== undefined) targetRange.format.font.bold = format.bold;
      if (format.italic !== undefined) targetRange.format.font.italic = format.italic;
      if (format.fontColor) targetRange.format.font.color = format.fontColor;
      if (format.fillColor) targetRange.format.fill.color = format.fillColor;
    }
    
    await context.sync();
  }
  
  /**
   * Undo a mergeCells operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoMergeCells(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo mergeCells: missing affected ranges');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Unmerge the cells
    targetRange.unmerge();
    
    // Restore the values if available
    if (action.beforeState.values) {
      targetRange.values = action.beforeState.values;
    }
    
    await context.sync();
  }
  
  /**
   * Undo an unmergeCells operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoUnmergeCells(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo unmergeCells: missing affected ranges');
      return;
    }
    
    const { sheetName, range } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const targetRange = worksheet.getRange(range);
    
    // Merge the cells again
    targetRange.merge();
    
    // Restore the value if available
    if (action.beforeState.values && action.beforeState.values.length > 0 && action.beforeState.values[0].length > 0) {
      targetRange.values = [[action.beforeState.values[0][0]]];
    }
    
    await context.sync();
  }
  
  /**
   * Undo a createChart operation
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoCreateChart(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo createChart: missing affected ranges');
      return;
    }
    
    const { sheetName } = action.affectedRanges[0];
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    
    // Load all charts
    const charts = worksheet.charts;
    charts.load('items');
    await context.sync();
    
    // Delete the last chart (assuming it's the one we created)
    if (charts.items.length > 0) {
      charts.items[charts.items.length - 1].delete();
    }
    
    await context.sync();
  }
  
  /**
   * Undo a composite operation by undoing each sub-operation in reverse order
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoCompositeOperation(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    // For composite operations, we need to extract the sub-operations and undo them
    const operation = action.operation;
    
    let subOperations: ExcelOperation[] = [];
    
    if ('operations' in operation && Array.isArray(operation.operations)) {
      subOperations = operation.operations;
    } else if ('subOperations' in operation && Array.isArray(operation.subOperations)) {
      subOperations = operation.subOperations;
    }
    
    if (subOperations.length === 0) {
      console.warn('Cannot undo composite operation: no sub-operations found');
      return;
    }
    
    // Create a mock action for each sub-operation
    // We'll use the affected ranges and before state from the parent action
    for (let i = subOperations.length - 1; i >= 0; i--) {
      const subOp = subOperations[i];
      
      // Create a mock action for this sub-operation
      const mockAction: WorkbookAction = {
        id: action.id + '-' + i,
        workbookId: action.workbookId,
        timestamp: action.timestamp,
        type: action.type,
        operation: subOp,
        description: `Sub-operation ${i} of ${action.description}`,
        affectedRanges: action.affectedRanges,
        beforeState: action.beforeState
      };
      
      // Undo this sub-operation
      await this.undoAction(context, mockAction);
    }
  }
  
  /**
   * Generic undo handler for operations without a specific handler
   * @param context The Excel context
   * @param action The action to undo
   */
  private async undoGeneric(context: Excel.RequestContext, action: WorkbookAction): Promise<void> {
    if (!action.affectedRanges.length) {
      console.warn('Cannot undo generic operation: missing affected ranges');
      return;
    }
    
    // For generic operations, we'll just try to restore values and formatting
    for (const { sheetName, range, type } of action.affectedRanges) {
      if (type === 'sheet' || !range) {
        continue; // Skip sheet-level operations or ranges without a reference
      }
      
      try {
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        const targetRange = worksheet.getRange(range);
        
        // Restore values and formulas if available
        if (action.beforeState.values) targetRange.values = action.beforeState.values;
        if (action.beforeState.formulas) targetRange.formulas = action.beforeState.formulas;
        
        // Restore formatting if available
        if (action.beforeState.formats && action.beforeState.formats.length > 0) {
          const format = action.beforeState.formats[0];
          
          if (format.fontFamily) targetRange.format.font.name = format.fontFamily;
          if (format.fontSize) targetRange.format.font.size = format.fontSize;
          if (format.bold !== undefined) targetRange.format.font.bold = format.bold;
          if (format.italic !== undefined) targetRange.format.font.italic = format.italic;
          if (format.fontColor) targetRange.format.font.color = format.fontColor;
          if (format.fillColor) targetRange.format.fill.color = format.fillColor;
        }
      } catch (error) {
        console.warn(`Error restoring range ${range} in sheet ${sheetName}:`, error);
      }
    }
    
    await context.sync();
  }
}
