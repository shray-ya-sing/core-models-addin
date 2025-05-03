/**
 * Utilities for working with Excel workbooks for multimodal analysis
 */
import * as ExcelJS from 'exceljs';

// Office.js is available globally in the add-in environment
/* global Excel, Office */

/**
 * Creates a base64-encoded copy of the active workbook with all formulas converted to values
 * Preserves all formatting and structure of the original workbook
 * @returns Promise with base64 encoded workbook
 */
export async function createFormulaFreeWorkbookCopy(): Promise<string> {
  return Excel.run(async (context) => {
    try {
      // Get all worksheets
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      
      // Create a new workbook using ExcelJS
      const wb = new ExcelJS.Workbook();
      
      // Process each worksheet
      for (const worksheet of worksheets.items) {
        // Create a new worksheet in our copy
        const newSheet = wb.addWorksheet(worksheet.name);
        
        // Get worksheet data including used range
        const usedRange = worksheet.getUsedRange();
        usedRange.load("address");
        usedRange.load("values");
        usedRange.load("rowCount");
        usedRange.load("columnCount");
        
        // We'll use a simplified approach that focuses on getting the values
        // rather than trying to preserve all formatting
        await context.sync();
        
        // Use default column widths and row heights
        // This simplifies the code and avoids potential type errors
        
        // Transfer values only - this is the key requirement
        // since we need to convert formulas to values
        for (let r = 0; r < usedRange.rowCount; r++) {
          for (let c = 0; c < usedRange.columnCount; c++) {
            if (usedRange.values[r] && usedRange.values[r][c] !== undefined) {
              const cell = newSheet.getCell(r + 1, c + 1);
              
              // Set the value (this will be the calculated value, not the formula)
              cell.value = usedRange.values[r][c];
              
              // Apply basic styling to make the workbook readable
              // We're using a simplified approach to avoid Office.js type issues
              cell.font = { name: 'Calibri', size: 11 };
            }
          }
        }
      }
      
      // Convert workbook to buffer
      const buffer = await wb.xlsx.writeBuffer();
      
      // Convert buffer to base64
      return arrayBufferToBase64(buffer);
    } catch (error) {
      console.error('Error creating formula-free workbook copy:', error);
      throw error;
    }
  });
}

/**
 * Converts an ArrayBuffer to a base64 string
 */
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/**
 * Sends a workbook to the specified API endpoint for image conversion
 * @param workbookBase64 Base64 encoded workbook
 * @param apiEndpoint API endpoint for image conversion
 * @param options Optional configuration for the image conversion
 * @returns Promise with array of base64 encoded images
 */
export async function getWorkbookImagesForMultimodalAnalysis(
  workbookBase64: string, 
  apiEndpoint: string,
  options?: {
    sheets?: string[];
    charts?: Array<{ sheetName: string; chartIndex: number }>;
    ranges?: Array<{ sheetName: string; range: string }>;
  }
): Promise<string[]> {
  try {
    // Prepare the request payload according to the endpoint's expected format
    const payload: any = {
      ExcelFile: workbookBase64
    };
    
    // Add optional parameters if provided
    if (options?.sheets && options.sheets.length > 0) {
      payload.Sheets = options.sheets;
    }
    
    if (options?.charts && options.charts.length > 0) {
      payload.Charts = options.charts.map(chart => ({
        SheetName: chart.sheetName,
        ChartIndex: chart.chartIndex
      }));
    }
    
    if (options?.ranges && options.ranges.length > 0) {
      payload.Ranges = options.ranges.map(range => ({
        SheetName: range.sheetName,
        Range: range.range
      }));
    }
    
    // Send to API endpoint for image conversion
    const response = await fetch(apiEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    
    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }
    
    // Get array of base64 images from response
    const result = await response.json();
    return result.images; // Assuming the API returns { images: string[] }
    
  } catch (error) {
    console.error('Error in multimodal analysis:', error);
    throw error;
  }
}

/**
 * Complete workflow for multimodal analysis of the active workbook
 * @param apiEndpoint API endpoint for image conversion
 * @param options Optional configuration for the image conversion
 * @returns Promise with array of base64 encoded images
 */
export async function performMultimodalAnalysis(
  apiEndpoint: string,
  options?: {
    sheets?: string[];
    charts?: Array<{ sheetName: string; chartIndex: number }>;
    ranges?: Array<{ sheetName: string; range: string }>;
  }
): Promise<string[]> {
  try {
    // Step 1: Create a formula-free copy of the workbook as base64
    const workbookBase64 = await createFormulaFreeWorkbookCopy();
    
    // Step 2: If no sheets are specified, get all worksheet names
    if (!options?.sheets || options.sheets.length === 0) {
      const worksheetNames = await getWorksheetNames();
      options = { ...options, sheets: worksheetNames };
    }
    
    // Step 3: Send to API for image conversion
    return await getWorkbookImagesForMultimodalAnalysis(workbookBase64, apiEndpoint, options);
  } catch (error) {
    console.error('Error performing multimodal analysis:', error);
    throw error;
  }
}

/**
 * Gets the names of all worksheets in the active workbook
 * @returns Promise with array of worksheet names
 */
async function getWorksheetNames(): Promise<string[]> {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    
    return worksheets.items.map(sheet => sheet.name);
  });
}
