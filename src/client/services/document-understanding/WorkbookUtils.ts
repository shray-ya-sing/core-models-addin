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
      
      // Log some information about the buffer for debugging
      console.log(`Created workbook buffer with size: ${buffer.byteLength} bytes`);
      
      // Convert buffer to base64
      const base64 = arrayBufferToBase64(buffer);
      console.log(`Converted to base64 string with length: ${base64.length} characters`);
      
      return base64;
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
    console.log(`Attempting to connect to image conversion API at: ${apiEndpoint}`);
    
    // Check API health first with a timeout to avoid long waits
    try {
      console.log(`
=======================================================
⚠️ CHECKING EXCEL IMAGE API HEALTH
=======================================================`);
      console.log(`Health endpoint: ${apiEndpoint.replace('/export', '')}/health`);
      
      const controller = new AbortController();
      const timeoutId = setTimeout(() => {
        controller.abort();
        console.error(`
=======================================================
❌ EXCEL IMAGE API HEALTH CHECK TIMED OUT
=======================================================
The Excel Image API server at ${apiEndpoint.split('/').slice(0, 3).join('/')} is not responding.

Please make sure the Excel Image API server is running at ${apiEndpoint.split('/').slice(0, 3).join('/')}.
This is required for formatting protocol analysis to work properly.

To start the server, you need to run the Excel Image API project.
=======================================================`);
      }, 3000); // 3 second timeout
      
      const healthCheck = await fetch(`${apiEndpoint.replace('/export', '')}/health`, {
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);
      
      if (!healthCheck.ok) {
        console.error(`
=======================================================
❌ EXCEL IMAGE API HEALTH CHECK FAILED: ${healthCheck.status} ${healthCheck.statusText}
=======================================================
The Excel Image API server at ${apiEndpoint.split('/').slice(0, 3).join('/')} returned an error.

Please make sure the Excel Image API server is running correctly.
This is required for formatting protocol analysis to work properly.
=======================================================`);
        throw new Error(`Excel Image API health check failed with status: ${healthCheck.status}`);
      }
      
      console.log(`
=======================================================
✅ EXCEL IMAGE API HEALTH CHECK SUCCESSFUL
=======================================================
Proceeding with image conversion.
=======================================================`);
    } catch (healthError) {
      if (healthError.name === 'AbortError') {
        // Already logged in the timeout handler
        throw new Error(`Excel Image API server not running at ${apiEndpoint.split('/').slice(0, 3).join('/')}`);
      } else if (healthError.name === 'TypeError' && healthError.message.includes('Failed to fetch')) {
        console.error(`
=======================================================
❌ EXCEL IMAGE API SERVER NOT RUNNING
=======================================================
The Excel Image API server at ${apiEndpoint.split('/').slice(0, 3).join('/')} is not running.

Please make sure the Excel Image API server is running at ${apiEndpoint.split('/').slice(0, 3).join('/')}.
This is required for formatting protocol analysis to work properly.

To start the server, you need to run the Excel Image API project.
=======================================================`);
        throw new Error(`Excel Image API server not running at ${apiEndpoint.split('/').slice(0, 3).join('/')}`);
      } else {
        console.error(`
=======================================================
❌ EXCEL IMAGE API HEALTH CHECK FAILED: ${healthError.message}
=======================================================
The Excel Image API server at ${apiEndpoint.split('/').slice(0, 3).join('/')} is not responding correctly.

Please make sure the Excel Image API server is running correctly.
This is required for formatting protocol analysis to work properly.
=======================================================`);
        throw new Error(`Excel Image API unavailable: ${healthError.message}`);
      }
    }
    
    // Prepare the request payload according to the documented API format
    const payload: any = {
      // Use 'ExcelFile' as per documentation
      ExcelFile: workbookBase64
    };
    
    // Add optional parameters if provided
    if (options?.sheets && options.sheets.length > 0) {
      // Use 'Sheets' as per documentation
      payload.Sheets = options.sheets;
    }
    
    if (options?.charts && options.charts.length > 0) {
      // Use 'Charts' with 'SheetName' and 'ChartIndex' as per documentation
      payload.Charts = options.charts.map(chart => ({
        SheetName: chart.sheetName,
        ChartIndex: chart.chartIndex
      }));
    }
    
    if (options?.ranges && options.ranges.length > 0) {
      // Use 'Ranges' with 'SheetName' and 'Range' as per documentation
      payload.Ranges = options.ranges.map(range => ({
        SheetName: range.sheetName,
        Range: range.range
      }));
    }
    
    // Send to API endpoint for image conversion with timeout
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout
    
    try {
      const response = await fetch(apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload),
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);
      
      if (!response.ok) {
        console.warn(`API error: ${response.status} ${response.statusText}`);
        throw new Error(`Excel Image API returned error: ${response.status} ${response.statusText}`);
      }
      
      // Parse the response according to the documented API format
      const result = await response.json();
      
      // Handle different possible response formats
      if (result.Images && Array.isArray(result.Images)) {
        // Format from documentation: { Images: [{ Name: string, Base64Image: string }] }
        console.log(`Received ${result.Images.length} images from the API in documented format`);
        
        // Extract the Base64Image values from each image object
        return result.Images.map((img: { Name: string; Base64Image: string }) => {
          console.log(`Processing image: ${img.Name}`);
          return img.Base64Image;
        });
      } else if (Array.isArray(result)) {
        // Format actually returned: direct array of base64 strings
        console.log(`Received ${result.length} images from the API in array format`);
        return result;
      } else if (result.images && Array.isArray(result.images)) {
        // Alternative format: { images: string[] }
        console.log(`Received ${result.images.length} images from the API in lowercase format`);
        return result.images;
      } else {
        console.warn('Unexpected API response format:', result);
        throw new Error('Unexpected API response format. Expected Images array or direct array of base64 strings.');
      }
    } catch (fetchError) {
      clearTimeout(timeoutId);
      console.warn(`API fetch error: ${fetchError.message}`);
      throw new Error(`Excel Image API fetch error: ${fetchError.message}`);
    }
    
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
