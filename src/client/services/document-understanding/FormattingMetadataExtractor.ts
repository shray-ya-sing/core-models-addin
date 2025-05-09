/**
 * Extractor for Excel workbook formatting metadata
 */
// Office.js is available globally in the Excel add-in environment
declare const Excel: any;
import { 
  CellFormatting, 
  OriginalFill, 
  SheetFormattingMetadata, 
  ThemeColors, 
  WorkbookFormattingMetadata 
} from './FormattingModels';

/**
 * Class for extracting formatting metadata from Excel workbooks
 */
export class FormattingMetadataExtractor {
  /**
   * Extracts formatting metadata from the active workbook
   * @returns Promise with workbook formatting metadata
   */
  public async extractFormattingMetadata(): Promise<WorkbookFormattingMetadata> {
    return Excel.run(async (context) => {
      try {
        // Get the workbook object
        const workbook = context.workbook;
        
        // Get all worksheets
        const worksheets = workbook.worksheets;
        worksheets.load("items/name");
        
        await context.sync();
        
        // Use default theme colors since we can't reliably access the theme
        const themeColors: ThemeColors = {
          background1: "#FFFFFF",
          background2: "#F2F2F2",
          text1: "#000000",
          text2: "#666666",
          accent1: "#4472C4",
          accent2: "#ED7D31",
          accent3: "#A5A5A5",
          accent4: "#FFC000",
          accent5: "#5B9BD5",
          accent6: "#70AD47",
          hyperlink: "#0563C1",
          followedHyperlink: "#954F72"
        };
        
        // Array to store sheet metadata
        const sheetsMetadata: SheetFormattingMetadata[] = [];
        
        // Process each worksheet
        for (const worksheet of worksheets.items) {
          // Get used range
          const usedRange = worksheet.getUsedRange();
          usedRange.load("address");
          
          // Get tables in the worksheet
          const tables = worksheet.tables;
          tables.load("items/name");
          
          // Get charts in the worksheet
          const charts = worksheet.charts;
          charts.load("items/name, items/chartType");
          
          await context.sync();
          
          // Skip if used range is null (empty worksheet)
          if (!usedRange.address) {
            continue;
          }
          
          // Get all cells in the used range
          const range = usedRange;
          
          // Load basic properties
          range.load("address, rowCount, columnCount");
          await context.sync();
          
          // Create a record to store cell formatting
          const cellFormatting: Record<string, CellFormatting> = {};
          
          // Store original fill colors to restore later
          const originalFills: Record<string, OriginalFill> = {};
          
          // Sample cells for formatting analysis
          // For large worksheets, we'll sample a subset of cells
          const rowCount = range.rowCount;
          const columnCount = range.columnCount;
          
          // Determine sampling rate based on worksheet size
          const rowSamplingRate = Math.max(1, Math.floor(rowCount / 20));
          const columnSamplingRate = Math.max(1, Math.floor(columnCount / 10));
          
          // Cells to load and process
          const cellsToLoad: Excel.Range[] = [];
          
          // Sample cells from the worksheet
          for (let row = 0; row < rowCount; row += rowSamplingRate) {
            for (let col = 0; col < columnCount; col += columnSamplingRate) {
              const cell = range.getCell(row, col);
              
              // Load formatting properties
              cell.load("address, format/fill/color, format/font/name, format/font/bold, format/font/color, numberFormat");
              cellsToLoad.push(cell);
            }
          }
          
          // Load all the selected cells
          await context.sync();
          
          // Initialize originalFills with loaded cell addresses
          for (const cell of cellsToLoad) {
            // Now it's safe to access cell.address since we've called context.sync()
            originalFills[cell.address] = {
              color: cell.format.fill.color,
              hasFill: !!cell.format.fill.color
            };
            
            // Load the cell value as text instead of using values
            // We'll convert this to string later
            cell.load("text");
          }
          
          // Sync again to load all the formatting properties
          await context.sync();
          
          // Process the cell formatting data and populate the cellFormatting object
          for (const cell of cellsToLoad) {
            // Extract the cell formatting data
            const cellAddress = cell.address;
            
            // Safely convert color values to strings
            const fillColor = cell.format.fill.color ? 
                            (typeof cell.format.fill.color === 'string' ? 
                             cell.format.fill.color : JSON.stringify(cell.format.fill.color)) : '';
            
            const fontName = cell.format.font.name || '';
            const fontBold = cell.format.font.bold || false;
            
            // Safely convert font color to string
            const fontColor = cell.format.font.color ? 
                            (typeof cell.format.font.color === 'string' ? 
                             cell.format.font.color : JSON.stringify(cell.format.font.color)) : '';
            const numberFormat = cell.numberFormat || '';
            // We need to properly type the cellValue to match the CellFormatting interface
            // which accepts string | any[] | any[][] for the value property
            let cellValue: string | any[] | any[][] = '';
            
            try {
              // Explicitly handle the cell.text property as any to avoid type errors
              const cellText: any = cell.text;
              
              if (cellText !== undefined && cellText !== null) {
                if (typeof cellText === 'string') {
                  // If it's already a string, use it directly
                  cellValue = cellText;
                } else if (Array.isArray(cellText)) {
                  // For arrays, keep the array structure
                  // This matches the CellFormatting interface which accepts string | any[][]
                  cellValue = cellText;
                } else {
                  // For any other type, convert to string
                  cellValue = String(cellText);
                }
              }
            } catch (error) {
              console.error('Error processing cell text:', error);
              // Use empty string as fallback
              cellValue = '';
            }
            
            // Store the cell formatting data
            cellFormatting[cellAddress] = {
              fillColor,
              fontName,
              fontBold,
              fontColor
            };
            
            // Restore the original fill color
            const originalFill = originalFills[cellAddress];
            
            if (!originalFill || !originalFill.hasFill) {
              // If there was no original fill, clear it
              cell.format.fill.clear();
            } else if (originalFill.color) {
              // Restore the original color as a string
              cell.format.fill.color = originalFill.color;
            }
          }
          
          // Apply the changes to restore original fill colors
          await context.sync();
          
          // Process tables
          const tablesData = [];
          for (const table of tables.items) {
            // Load the table range
            const tableRange = table.getRange();
            tableRange.load("address");
            await context.sync();
            
            tablesData.push({
              name: table.name,
              range: tableRange.address
            });
          }
          
          // Process charts
          const chartsData: Record<string, { chartType: string }> = {};
          charts.items.forEach(chart => {
            chartsData[chart.name] = {
              chartType: chart.chartType
            };
          });
          
          // Add sheet metadata
          sheetsMetadata.push({
            name: worksheet.name,
            cells: cellFormatting,
            tables: tablesData,
            charts: chartsData,
            commonFormats: {} // Will be populated by the LLM
          });
        }
        
        return {
          themeColors,
          sheets: sheetsMetadata
        };
      } catch (error) {
        console.error('Error extracting formatting metadata:', error);
        throw error;
      }
    });
  }
}
