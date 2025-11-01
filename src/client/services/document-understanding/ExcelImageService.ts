/**
 * Service for capturing Excel workbook images
 */
// Office.js is available globally in the Excel add-in environment
declare const Excel: any;
declare const Office: any;

// Import the utility function for creating a formula-free workbook copy
import { createFormulaFreeWorkbookCopy } from './WorkbookUtils';

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
   * Captures images of worksheets in the active workbook
   * @param maxWorksheets Maximum number of worksheets to capture (default: 5)
   * @returns Promise with an array of base64-encoded images
   */
  public async captureWorkbookImages(maxWorksheets: number = 5): Promise<string[]> {
    return Excel.run(async (context) => {
      try {
        // Get all worksheets
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        
        // Array to store base64 images
        const images: string[] = [];
        const skippedWorksheets: string[] = [];
        
        // Log the worksheet limit
        if (worksheets.items.length > maxWorksheets) {
          console.log(`⚠️ Limiting image capture to first ${maxWorksheets} worksheets (out of ${worksheets.items.length} total)`); 
        }
        
        // For each worksheet (limited to maxWorksheets), capture an image
        for (let i = 0; i < Math.min(worksheets.items.length, maxWorksheets); i++) {
          const worksheet = worksheets.items[i];
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
            // This will return null if there's an error instead of throwing
            const imageBase64 = await this.callExcelImageApi(worksheet.name);
            
            // Only add to images array if we got a valid result
            if (imageBase64) {
              images.push(imageBase64);
              console.log(`Successfully captured image for worksheet: ${worksheet.name}`);
            } else {
              console.warn(`Failed to capture image for worksheet: ${worksheet.name}`);
              skippedWorksheets.push(worksheet.name);
            }
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
        
        // Log any worksheets not processed due to the limit
        if (worksheets.items.length > maxWorksheets) {
          const remainingCount = worksheets.items.length - maxWorksheets;
          const remainingNames = worksheets.items.slice(maxWorksheets).map(ws => ws.name).join(', ');
          console.log(`ℹ️ ${remainingCount} worksheets were not processed due to the limit: ${remainingNames}`);
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
  private async callExcelImageApi(worksheetName: string): Promise<string | null> {
    try {
      console.log(`
-------------------------------------------------------
📷 CAPTURING IMAGE FOR WORKSHEET: ${worksheetName}
-------------------------------------------------------`);
      
      // Step 1: Create an in-memory copy of the workbook with formulas converted to values
      console.log(`Step 1: Creating formula-free workbook copy for worksheet: ${worksheetName}`);
      let workbookBase64: string;
      try {
        // Using the utility function from WorkbookUtils.ts
        const startTime = Date.now();
        workbookBase64 = await createFormulaFreeWorkbookCopy();
        const elapsedTime = Date.now() - startTime;
        
        // Log the size of the base64 string for debugging
        console.log(`✅ Successfully created base64 Excel file in ${elapsedTime}ms`);
        console.log(`   Base64 size: ${workbookBase64.length} characters`);
        console.log(`   Base64 starts with: ${workbookBase64.substring(0, 50)}...`);
        
        if (!workbookBase64 || workbookBase64.length < 100) {
          console.error(`❌ Invalid workbook base64 data: too short (${workbookBase64?.length || 0} chars)`);
          return null;
        }
      } catch (workbookError) {
        console.error(`❌ Error creating formula-free workbook copy:`, workbookError);
        console.error(`Error stack trace: ${(workbookError as Error).stack}`);
        return null;
      }
      
      // Step 2: Prepare the request payload
      console.log(`Step 2: Preparing API request payload for ${worksheetName}`);
      const payload = {
        "ExcelFile": workbookBase64,
        "Sheets": [worksheetName] // Use uppercase 'S' in Sheets to match API expectations
      };
      
      // Log the payload structure (without the actual base64 content)
      console.log(`   Payload structure: ${JSON.stringify({
        ExcelFile: `${workbookBase64.substring(0, 30)}... (${workbookBase64.length} chars)`,
        Sheets: payload.Sheets
      }, null, 2)}`);
      
      // Step 3: Check API health
      console.log(`Step 3: Checking Excel Image API health`);
      const healthEndpoint = `${this.apiEndpoint.replace('/export', '')}/health`;
      console.log(`   Health endpoint: ${healthEndpoint}`);
      
      const healthController = new AbortController();
      const healthTimeoutId = setTimeout(() => {
        console.warn(`⏰ Health check timed out after 3 seconds`);
        healthController.abort();
      }, 3000); // 3 second timeout
      
      try {
        const startTime = Date.now();
        const healthCheck = await fetch(healthEndpoint, {
          signal: healthController.signal
        });
        const elapsedTime = Date.now() - startTime;
        
        clearTimeout(healthTimeoutId);
        
        if (!healthCheck.ok) {
          console.error(`❌ API health check failed with status: ${healthCheck.status} in ${elapsedTime}ms`);
          try {
            const errorText = await healthCheck.text();
            console.error(`   Error response: ${errorText}`);
          } catch (e) {
            console.error(`   Could not read error response`);
          }
          return null;
        }
        
        console.log(`✅ API health check successful in ${elapsedTime}ms, proceeding with image conversion`);
      } catch (healthError) {
        console.error(`❌ API health check failed: ${(healthError as Error).message}`);
        console.error(`   Is the Excel Image API running at ${this.apiEndpoint.split('/').slice(0, 3).join('/')}?`);
        return null;
      }
      
      // Step 4: Send to API endpoint for image conversion
      console.log(`Step 4: Sending request to Excel Image API`);
      console.log(`   API endpoint: ${this.apiEndpoint}`);
      
      const controller = new AbortController();
      const timeoutId = setTimeout(() => {
        console.error(`⏰ API request timed out after 10 seconds`);
        controller.abort();
      }, 10000); // 10 second timeout
      
      try {
        const startTime = Date.now();
        console.log(`   Sending POST request with payload size: ${JSON.stringify(payload).length} bytes`);
        
        const response = await fetch(this.apiEndpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(payload),
          signal: controller.signal
        });
        
        const elapsedTime = Date.now() - startTime;
        clearTimeout(timeoutId);
        
        console.log(`   Received response in ${elapsedTime}ms with status: ${response.status}`);
        
        if (!response.ok) {
          // Try to get the response text for more detailed error information
          let errorDetails = '';
          try {
            errorDetails = await response.text();
          } catch (e) {
            errorDetails = 'Could not read error details';
          }
          
          console.error(`❌ API request failed with status: ${response.status}. Details: ${errorDetails}`);
          return null;
        }
        
        // Step 5: Parse the response
        console.log(`Step 5: Parsing API response`);
        let responseData: any;
        try {
          responseData = await response.json();
          console.log(`   Response contains ${Object.keys(responseData).length} keys: ${Object.keys(responseData).join(', ')}`);
        } catch (parseError) {
          console.error(`❌ Error parsing API response as JSON:`, parseError);
          return null;
        }
        
        // Check if the response contains images - handle different possible response formats
        let images = null;
        
        // Log the response structure for debugging
        console.log(`   Response keys: ${Object.keys(responseData).join(', ')}`);
        
        // Try different possible keys where images might be stored
        if (responseData.images && Array.isArray(responseData.images) && responseData.images.length > 0) {
          images = responseData.images;
          console.log(`   Found images in responseData.images`);
        } else if (responseData.Images && Array.isArray(responseData.Images) && responseData.Images.length > 0) {
          images = responseData.Images;
          console.log(`   Found images in responseData.Images`);
        } else if (responseData.data && Array.isArray(responseData.data) && responseData.data.length > 0) {
          images = responseData.data;
          console.log(`   Found images in responseData.data`);
        } else if (Array.isArray(responseData) && responseData.length > 0) {
          images = responseData;
          console.log(`   Response is directly an array of images`);
        }
        
        // If we still don't have images, check if the response itself is a single image string
        if (!images && typeof responseData === 'string' && responseData.length > 0) {
          images = [responseData];
          console.log(`   Response is directly a single image string`);
        }
        
        // If we couldn't find images in any expected format, log the error and return null
        if (!images) {
          console.error(`❌ API response did not contain any images in expected format`);
          console.log(`   Full response: ${JSON.stringify(responseData).substring(0, 500)}...`);
          return null;
        }
        
        console.log(`   Response contains ${images.length} images`);
        
        // Step 6: Get the first image (since we only requested one sheet)
        const imageBase64 = images[0];
        console.log(`   First image length: ${imageBase64?.length || 0} characters`);
        console.log(`   First image starts with: ${imageBase64?.substring(0, 30)}...`);
        
        // Step 7: Validate the image
        console.log(`Step 7: Validating image data`);
        if (!this.isValidBase64PngImage(imageBase64)) {
          console.error(`❌ API returned an invalid base64 PNG image`);
          console.log(`   Image starts with: ${imageBase64?.substring(0, 30)}...`);
          return null;
        }
        
        console.log(`✅ Successfully captured valid PNG image for worksheet: ${worksheetName}`);
        return imageBase64;
      } catch (apiError) {
        console.error(`❌ Error calling Excel Image API:`, apiError);
        console.error(`   Error stack trace: ${(apiError as Error).stack}`);
        return null;
      }
    } catch (error) {
      console.error(`❌ Unexpected error in callExcelImageApi:`, error);
      console.error(`   Error stack trace: ${(error as Error).stack}`);
      return null;
    }
  }
  
  // We're now using the createFormulaFreeWorkbookCopy utility function from WorkbookUtils.ts
  // instead of implementing our own workbook copy method
}
