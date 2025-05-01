// src/client/services/ClientExcelCommandInterpreter.ts
// Interprets and executes Excel operations using the Office.js API

import { 
  ExcelOperation, 
  ExcelOperationType,
  ExcelCommandPlan
} from '../models/ExcelOperationModels';
import * as ExcelUtils from '../utils/ExcelUtils';

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
          
        // Worksheet settings operations
        case ExcelOperationType.SET_GRIDLINES:
          await this.executeSetGridlines(context, operation);
          break;
          
        case ExcelOperationType.SET_HEADERS:
          await this.executeSetHeaders(context, operation);
          break;
          
        case ExcelOperationType.SET_ZOOM:
          await this.executeSetZoom(context, operation);
          break;
          
        case ExcelOperationType.SET_FREEZE_PANES:
          await this.executeSetFreezePanes(context, operation);
          break;
          
        case ExcelOperationType.SET_VISIBLE:
          await this.executeSetVisible(context, operation);
          break;
          
        case ExcelOperationType.SET_ACTIVE_SHEET:
          await this.executeSetActiveSheet(context, operation);
          break;
          
        // Print settings operations
        case ExcelOperationType.SET_PRINT_AREA:
          await this.executeSetPrintArea(context, operation);
          break;
          
        case ExcelOperationType.SET_PRINT_ORIENTATION:
          await this.executeSetPrintOrientation(context, operation);
          break;
          
        case ExcelOperationType.SET_PRINT_MARGINS:
          await this.executeSetPrintMargins(context, operation);
          break;
          
        // Chart formatting operations
        case ExcelOperationType.FORMAT_CHART:
          await this.executeFormatChart(context, operation);
          break;
          
        case ExcelOperationType.FORMAT_CHART_AXIS:
          await this.executeFormatChartAxis(context, operation);
          break;
          
        case ExcelOperationType.FORMAT_CHART_SERIES:
          await this.executeFormatChartSeries(context, operation);
          break;
          
        // Complex operations
        case ExcelOperationType.COMPOSITE_OPERATION:
          await this.executeCompositeOperation(context, operation);
          break;
          
        case ExcelOperationType.BATCH_OPERATION:
          await this.executeBatchOperation(context, operation);
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

  /**
   * Execute a SET_GRIDLINES operation
   */
  private async executeSetGridlines(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    worksheet.showGridlines = op.display;
  }

  /**
   * Execute a SET_HEADERS operation
   */
  private async executeSetHeaders(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    worksheet.showHeadings = op.display;
  }

  /**
   * Execute a SET_ZOOM operation
   */
  private async executeSetZoom(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    // Set zoom using pageLayout.zoom with a scale property
    worksheet.pageLayout.zoom = { scale: op.zoomLevel };
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
   * Execute a SET_VISIBLE operation
   */
  private async executeSetVisible(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    if (op.visible) {
      worksheet.visibility = Excel.SheetVisibility.visible;
    } else {
      worksheet.visibility = Excel.SheetVisibility.hidden;
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
   * Execute a SET_PRINT_AREA operation
   */
  private async executeSetPrintArea(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    worksheet.pageLayout.setPrintArea(op.range);
  }

  /**
   * Execute a SET_PRINT_ORIENTATION operation
   */
  private async executeSetPrintOrientation(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    
    worksheet.pageLayout.orientation = 
      op.orientation === 'landscape' ? Excel.PageOrientation.landscape : Excel.PageOrientation.portrait;
  }

  /**
   * Execute a SET_PRINT_MARGINS operation
   */
  private async executeSetPrintMargins(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
    const op = operation as any;
    const worksheet = context.workbook.worksheets.getItem(op.sheet);
    const pageLayout = worksheet.pageLayout;
    
    if (op.hasOwnProperty('top')) pageLayout.topMargin = op.top;
    if (op.hasOwnProperty('right')) pageLayout.rightMargin = op.right;
    if (op.hasOwnProperty('bottom')) pageLayout.bottomMargin = op.bottom;
    if (op.hasOwnProperty('left')) pageLayout.leftMargin = op.left;
    if (op.hasOwnProperty('header')) pageLayout.headerMargin = op.header;
    if (op.hasOwnProperty('footer')) pageLayout.footerMargin = op.footer;
  }

  /**
   * Execute a FORMAT_CHART operation
   */
  private async executeFormatChart(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
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
    
    // Apply formatting
    if (op.hasOwnProperty('title')) {
      chart.title.text = op.title;
    }
    
    if (op.hasOwnProperty('hasLegend')) {
      chart.legend.visible = op.hasLegend;
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
    
    if (op.hasOwnProperty('height')) {
      chart.height = op.height;
    }
    
    if (op.hasOwnProperty('width')) {
      chart.width = op.width;
    }
    
    if (op.hasOwnProperty('style')) {
      chart.style = op.style;
    }
  }

  /**
   * Execute a FORMAT_CHART_AXIS operation
   */
  private async executeFormatChartAxis(context: Excel.RequestContext, operation: ExcelOperation): Promise<void> {
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
    
    // Get the axis
    let axis: Excel.ChartAxis;
    
    // Use the getItem method to get the specific axis by type and group
    if (op.axisType === 'category') {
      // For category axis
      axis = chart.axes.getItem(
        Excel.ChartAxisType.category, 
        op.axisGroup === 'secondary' ? Excel.ChartAxisGroup.secondary : Excel.ChartAxisGroup.primary
      );
    } else {
      // For value axis
      axis = chart.axes.getItem(
        Excel.ChartAxisType.value, 
        op.axisGroup === 'secondary' ? Excel.ChartAxisGroup.secondary : Excel.ChartAxisGroup.primary
      );
    }

    
    // Apply formatting
    if (op.hasOwnProperty('title') && op.hasOwnProperty('hasTitle')) {
      axis.title.text = op.title;
      axis.title.visible = op.hasTitle;
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
    
    if (op.hasOwnProperty('logScale')) {
      axis.scaleType = Excel.ChartAxisScaleType.logarithmic;
      axis.logBase   = op.logBase ?? 10;       // optional numeric base
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
}
