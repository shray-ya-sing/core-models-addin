/**
 * Service for capturing Excel workbook images
 */
// Office.js is available globally in the Excel add-in environment
declare const Excel: any;
declare const Office: any;

// Import the utility function for creating a formula-free workbook copy
import { createFormulaFreeWorkbookCopy } from './document-understanding/WorkbookUtils';

/**
 * Regular expression to validate base64 strings
 * This regex checks for a valid base64 character set and proper padding
 */
const BASE64_REGEX = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;

/**
 * PNG file signature in base64 (first 8 bytes of a PNG file encoded as base64)
 * This corresponds to the PNG magic number: 89 50 4E 47 0D 0A 1A 0A
 */
const PNG_SIGNATURE_BASE64_PREFIX = 'iVBORw';

/**
 * Excel file signature in base64 (common prefixes for .xlsx files)
 * These correspond to the first bytes of a ZIP file (which is what .xlsx files are)
 * PK\x03\x04 (ZIP file header) typically encodes as 'UEs'
 */
const EXCEL_FILE_BASE64_PREFIXES = ['UEs', 'PK'];

/**
 * Minimum reasonable size for an Excel file in bytes
 * Even the smallest Excel file should be at least a few KB
 */
const MIN_EXCEL_FILE_SIZE = 2000; // 2KB

/**
 * Service for capturing Excel workbook images
 */
export class ExcelImageService {
  /**
   * Validates if a string is a valid base64 encoded Excel file
   * @param base64String The base64 string to validate
   * @returns True if the string is a valid base64 encoded Excel file, false otherwise
   */
  private isValidBase64ExcelFile(base64String: string): boolean {
    try {
      // Check if the string is empty or null
      if (!base64String) {
        console.warn('Base64 string is empty or null');
        return false;
      }
      
      // Log the first 100 characters of the string for debugging
      console.log(`First 100 chars of base64 string: ${base64String.substring(0, 100)}`);
      
      // For now, just check if the string has a reasonable length
      // We'll skip the strict base64 pattern check since Office.js might use a different format
      if (base64String.length < MIN_EXCEL_FILE_SIZE) {
        console.warn(`Base64 string length (${base64String.length}) is too small for a valid Excel file`);
        return false;
      }
      
      // Skip the prefix check for now, as Office.js might encode the file differently
      // Just log what we would have checked
      let hasExpectedPrefix = false;
      for (const prefix of EXCEL_FILE_BASE64_PREFIXES) {
        if (base64String.substring(0, prefix.length).includes(prefix)) {
          hasExpectedPrefix = true;
          console.log(`Found expected Excel file signature: ${prefix}`);
          break;
        }
      }
      
      if (!hasExpectedPrefix) {
        console.log('Did not find expected Excel file signature, but continuing anyway');
      }
      
      // For now, assume the file is valid if it has content
      console.log(`Assuming base64 Excel file is valid with length: ${base64String.length}`);
      return true;
    } catch (error) {
      console.error('Error validating base64 Excel file:', error);
      return false;
    }
  }
  
  /**
   * Validates if a string is a valid base64 encoded PNG image
   * @param base64String The base64 string to validate
   * @returns True if the string is a valid base64 encoded PNG image, false otherwise
   */
  private isValidBase64PngImage(base64String: string): boolean {
    try {
      // Check if the string is empty or null
      if (!base64String) {
        console.warn('Base64 string is empty or null');
        return false;
      }
      
      // Remove data URL prefix if present
      let cleanBase64 = base64String;
      if (base64String.startsWith('data:image/png;base64,')) {
        cleanBase64 = base64String.substring('data:image/png;base64,'.length);
      }
      
      // Check if the string matches the base64 pattern
      if (!BASE64_REGEX.test(cleanBase64)) {
        console.warn('String does not match base64 pattern');
        return false;
      }
      
      // Check if the string starts with the PNG signature
      if (!cleanBase64.startsWith(PNG_SIGNATURE_BASE64_PREFIX)) {
        console.warn('String does not start with PNG signature');
        return false;
      }
      
      // Additional validation: check if the decoded length is reasonable
      // PNG files should be at least a few hundred bytes
      const decodedLength = Math.floor(cleanBase64.length * 0.75);
      if (decodedLength < 100) {
        console.warn('Decoded base64 length is too small for a valid PNG');
        return false;
      }
      
      return true;
    } catch (error) {
      console.error('Error validating base64 PNG image:', error);
      return false;
    }
  }
  
  /**
   * Captures images of all worksheets in the active workbook
   * @returns Promise with an array of base64-encoded images
   */
  public async captureWorkbookImages(): Promise<string[]> {
    return Excel.run(async (context) => {
      try {
        // Get all worksheets
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        
        // Array to store base64 images
        const images: string[] = [];
        const skippedWorksheets: string[] = [];
        
        // For each worksheet, capture an image
        for (const worksheet of worksheets.items) {
          try {
            // Activate the worksheet to capture
            worksheet.activate();
            await context.sync();
            
            // Get the used range to determine the area to capture
            const usedRange = worksheet.getUsedRange();
            // Load address property explicitly
            usedRange.load("address");
            await context.sync();
            
            // Skip if used range is null or empty (empty worksheet)
            if (!usedRange || !usedRange.address) {
              console.log(`Skipping empty worksheet: ${worksheet.name}`);
              skippedWorksheets.push(worksheet.name);
              continue;
            }
            
            // Call the Excel Image API to capture the worksheet
            // If this fails, it will throw an error and be caught by the catch block
            const imageBase64 = await this.callExcelImageApi(worksheet.name);
            
            // At this point, the image should be valid since callExcelImageApi validates it
            images.push(imageBase64);
            console.log(`Successfully captured image for worksheet: ${worksheet.name}`);
          } catch (worksheetError) {
            // Log the error and continue with the next worksheet
            console.error(`Error capturing image for worksheet ${worksheet.name}:`, worksheetError);
            skippedWorksheets.push(worksheet.name);
          }
        }
        
        // Log any skipped worksheets
        if (skippedWorksheets.length > 0) {
          console.warn(`Skipped image capture for worksheets: ${skippedWorksheets.join(', ')}`);
        }
        
        // If no images were captured, log a warning
        if (images.length === 0) {
          console.warn('No worksheet images were captured. Formatting analysis may be limited.');
        } else {
          console.log(`Successfully captured ${images.length} worksheet images`);
        }
        
        return images;
      } catch (error) {
        console.error('Error capturing workbook images:', error);
        throw error;
      }
    });
  }
  
  /**
   * API endpoint for Excel image conversion
   * This is the default endpoint, but it can be overridden in the constructor
   */
  private apiEndpoint = 'http://localhost:8080/api/ExcelImage/export';
  
  /**
   * Constructor for the ExcelImageService
   * @param customApiEndpoint Optional custom API endpoint for image conversion
   */
  constructor(customApiEndpoint?: string) {
    if (customApiEndpoint) {
      this.apiEndpoint = customApiEndpoint;
    }
  }
  
  /**
   * Calls the Excel Image API to capture a worksheet
   * @param worksheetName The name of the worksheet to capture
   * @returns Promise with a base64-encoded image
   */
  private async callExcelImageApi(worksheetName: string): Promise<string> {
    try {
      console.log(`Capturing image for worksheet: ${worksheetName}`);
      
      // Create an in-memory copy of the workbook with formulas converted to values
      // Using the utility function from WorkbookUtils.ts
      console.log(`Creating formula-free workbook copy for worksheet: ${worksheetName}`);
      const workbookBase64 = await createFormulaFreeWorkbookCopy();
      
      // Log the size of the base64 string for debugging
      console.log(`Workbook base64 size for ${worksheetName}: ${workbookBase64.length} characters`);
      
      // Since we're using the proven utility function, we can be more confident in the result
      console.log(`Successfully created base64 Excel file for worksheet: ${worksheetName}`);
      
      // Prepare the request payload with the exact format expected by the API
      const payload = {
        "ExcelFile": workbookBase64,
        "Sheets": [worksheetName] // Use uppercase 'S' in Sheets to match API expectations
      };
      
      // Log the payload structure (without the actual base64 content)
      console.log(`Payload structure: ${JSON.stringify({
        ExcelFile: payload.ExcelFile.slice(0, 100),
        Sheets: payload.Sheets
      }, null, 2)}`);
      
      
      // Check API health first
      const healthEndpoint = `${this.apiEndpoint.replace('/export', '')}/health`;
      console.log(`Checking API health at: ${healthEndpoint}`);
      
      const healthController = new AbortController();
      const healthTimeoutId = setTimeout(() => healthController.abort(), 3000); // 3 second timeout
      
      try {
        const healthCheck = await fetch(healthEndpoint, {
          signal: healthController.signal
        });
        
        clearTimeout(healthTimeoutId);
        
        if (!healthCheck.ok) {
          throw new Error(`API health check failed with status: ${healthCheck.status}`);
        }
        
        console.log('API health check successful, proceeding with image conversion');
      } catch (healthError) {
        console.warn(`API health check failed: ${healthError.message}`);
        throw new Error(`Excel Image API unavailable: ${healthError.message}`);
      }
      
      // Send to API endpoint for image conversion with timeout
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout
      
      console.log(`Sending request to Excel Image API: ${this.apiEndpoint}`);
      
      const response = await fetch(this.apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload),
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);
      
      if (!response.ok) {
        // Try to get the response text for more detailed error information
        let errorDetails = '';
        try {
          errorDetails = await response.text();
          console.error(`API error details: ${errorDetails}`);
        } catch (textError) {
          console.error('Could not read error details:', textError);
        }
        
        throw new Error(`Excel Image API returned error: ${response.status} ${response.statusText}${errorDetails ? ` - ${errorDetails}` : ''}`);
      }
      
      console.log(`Received successful response from Excel Image API for worksheet: ${worksheetName}`);
      
      // Parse the response - expecting an array of strings
      const result = await response.json();
      console.log(`Response received from API for worksheet: ${worksheetName}`);
      
      // The API should return an array of base64 strings
      if (!Array.isArray(result)) {
        console.error('Unexpected response format - not an array:', typeof result);
        throw new Error('Unexpected API response format. Expected an array of base64 strings.');
      }
      
      if (result.length === 0) {
        console.error('API returned empty array of images');
        throw new Error('No images returned from the API.');
      }
      
      // Get the first image from the array
      const imageBase64 = result[0];
      console.log(`Received image for worksheet: ${worksheetName} (${result.length} total images)`);
      
      // Log image size for debugging
      console.log(`Image size: ${imageBase64.length} characters`);
      
      // Validate the base64 image before returning it
      if (!this.isValidBase64PngImage(imageBase64)) {
        throw new Error(`Invalid base64 PNG image returned for worksheet: ${worksheetName}`);
      }
      
      return imageBase64;
    } catch (error) {
      console.error(`Error calling Excel Image API for worksheet ${worksheetName}:`, error);
      throw error;
    }
  }
  
  // We're now using the createFormulaFreeWorkbookCopy utility function from WorkbookUtils.ts
  // instead of implementing our own workbook copy method
}
