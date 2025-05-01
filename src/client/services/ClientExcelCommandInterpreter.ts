// src/client/services/ClientExcelCommandInterpreter.ts
// Interprets and executes Excel operations using the Office.js API

import { 
  ExcelOperation, 
  ExcelOperationType,
  ExcelCommandPlan
} from '../models/ExcelOperationModels';

/**
 * Service that interprets and executes Excel operations using Office.js
 */
export class ClientExcelCommandInterpreter {
  /**
   * Execute a command plan with multiple operations
   * @param plan The Excel command plan to execute
   * @returns A promise that resolves when all operations are complete
   */
  public async executeCommandPlan(plan: ExcelCommandPlan): Promise<void> {
    console.log(`Executing command plan: ${plan.description}`);
    console.log(`Operations to execute: ${plan.operations.length}`);
    
    return this.executeOperations(plan.operations);
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
    
    try {
      await Excel.run(async (context) => {
        for (const operation of operations) {
          await this.executeOperation(context, operation);
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
   * Execute a single Excel operation
   * @param context The Excel context
   * @param operation The operation to execute
   */
  private async executeOperation(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    console.log(`Executing operation: ${operation.op}`);
    
    try {
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
          
        case ExcelOperationType.RENAME_SHEET:
          await this.executeRenameSheet(context, operation);
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
          
        default:
          console.warn(`Unsupported operation: ${(operation as any).op}`);
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
        case 'percentage':
          formatString = '0.00%';
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
        default:
          formatString = op.style;
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
      }
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
   * Execute a RENAME_SHEET operation
   */
  private async executeRenameSheet(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    context.workbook.worksheets.getItem(op.oldName).name = op.newName;
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
}
