import { Operation, OperationType } from '../models/CommandModels';
import { ClientWorkbookStateManager } from './ClientWorkbookStateManager';

/**
 * Client-side command executor for executing operations in Excel
 */
export class ClientCommandExecutor {
  private workbookStateManager: ClientWorkbookStateManager;

  constructor(workbookStateManager: ClientWorkbookStateManager) {
    this.workbookStateManager = workbookStateManager;
  }

  /**
   * Execute an operation in Excel
   * @param operation The operation to execute
   */
  public async executeOperation(operation: Operation): Promise<void> {
    try {
      switch (operation.type) {
        case OperationType.SetValue:
          await this.setValue(operation.target, operation.value);
          break;
        case OperationType.SetFormula:
          await this.setFormula(operation.target, operation.value);
          break;
        case OperationType.FormatCell:
          await this.formatCell(operation.target, operation.options);
          break;
        case OperationType.CreateSheet:
          await this.createSheet(operation.value);
          break;
        case OperationType.DeleteSheet:
          await this.deleteSheet(operation.target);
          break;
        case OperationType.RenameSheet:
          await this.renameSheet(operation.target, operation.value);
          break;
        case OperationType.CreateTable:
          await this.createTable(operation.target, operation.options);
          break;
        case OperationType.CreateChart:
          await this.createChart(operation.target, operation.value, operation.options);
          break;
        default:
          throw new Error(`Unsupported operation type: ${operation.type}`);
      }
    } catch (error) {
      console.error(`Error executing operation ${operation.type}:`, error);
      throw error;
    }
  }

  /**
   * Set a cell value
   * @param target The target cell or range
   * @param value The value to set
   */
  private async setValue(target: string, value: any): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(target);
        range.values = [[value]];
        await context.sync();
      });
    } catch (error) {
      console.error('Error setting value:', error);
      throw error;
    }
  }

  /**
   * Set a cell formula
   * @param target The target cell or range
   * @param formula The formula to set
   */
  private async setFormula(target: string, formula: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(target);
        range.formulas = [[formula]];
        await context.sync();
      });
    } catch (error) {
      console.error('Error setting formula:', error);
      throw error;
    }
  }

  /**
   * Format a cell
   * @param target The target cell or range
   * @param options Formatting options
   */
  private async formatCell(target: string, options: any): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(target);
        
        // Apply formatting options
        if (options.bold !== undefined) {
          range.format.font.bold = options.bold;
        }
        if (options.italic !== undefined) {
          range.format.font.italic = options.italic;
        }
        if (options.underline !== undefined) {
          range.format.font.underline = options.underline;
        }
        if (options.fontSize !== undefined) {
          range.format.font.size = options.fontSize;
        }
        if (options.fontColor !== undefined) {
          range.format.font.color = options.fontColor;
        }
        if (options.fillColor !== undefined) {
          range.format.fill.color = options.fillColor;
        }
        if (options.numberFormat !== undefined) {
          range.numberFormat = options.numberFormat;
        }
        if (options.horizontalAlignment !== undefined) {
          range.format.horizontalAlignment = options.horizontalAlignment;
        }
        if (options.verticalAlignment !== undefined) {
          range.format.verticalAlignment = options.verticalAlignment;
        }
        
        await context.sync();
      });
    } catch (error) {
      console.error('Error formatting cell:', error);
      throw error;
    }
  }

  /**
   * Create a new worksheet
   * @param name The name of the new worksheet
   */
  private async createSheet(name: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.add(name);
        sheet.activate();
        await context.sync();
      });
    } catch (error) {
      console.error('Error creating sheet:', error);
      throw error;
    }
  }

  /**
   * Delete a worksheet
   * @param name The name of the worksheet to delete
   */
  private async deleteSheet(name: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(name);
        sheet.delete();
        await context.sync();
      });
    } catch (error) {
      console.error('Error deleting sheet:', error);
      throw error;
    }
  }

  /**
   * Rename a worksheet
   * @param currentName The current name of the worksheet
   * @param newName The new name for the worksheet
   */
  private async renameSheet(currentName: string, newName: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(currentName);
        sheet.name = newName;
        await context.sync();
      });
    } catch (error) {
      console.error('Error renaming sheet:', error);
      throw error;
    }
  }

  /**
   * Create a table
   * @param range The range for the table
   * @param options Table options
   */
  private async createTable(range: string, options: any): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const tableRange = worksheet.getRange(range);
        const table = worksheet.tables.add(tableRange, options?.hasHeaders ?? true);
        
        if (options?.name) {
          table.name = options.name;
        }
        if (options?.style) {
          table.style = options.style;
        }
        
        await context.sync();
      });
    } catch (error) {
      console.error('Error creating table:', error);
      throw error;
    }
  }

  /**
   * Create a chart
   * @param range The data range for the chart
   * @param type The chart type
   * @param options Chart options
   */
  private async createChart(range: string, type: string, options: any): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const dataRange = worksheet.getRange(range);
        
        // Convert string type to Excel.ChartType
        // This handles the case where type comes as a string from the API
        const chartType = Excel.ChartType[type as keyof typeof Excel.ChartType] || type;
        
        // Create the chart
        const chart = worksheet.charts.add(chartType as Excel.ChartType, dataRange, 'Auto');
        
        // Set chart options
        if (options?.title) {
          chart.title.text = options.title;
        }
        if (options?.position) {
          chart.setPosition(options.position.top, options.position.left);
        }
        if (options?.width && options?.height) {
          chart.width = options.width;
          chart.height = options.height;
        }
        // Excel JS API uses 'series' instead of 'seriesBy'
        if (options?.seriesBy) {
          // @ts-ignore - Handle seriesBy option for backward compatibility
          chart.series = options.seriesBy;
        }
        
        await context.sync();
      });
    } catch (error) {
      console.error('Error creating chart:', error);
      throw error;
    }
  }
}
