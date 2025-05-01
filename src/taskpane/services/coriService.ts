/**
 * Service for communicating with the Cori backend server
 */

import config from "../../client/config";

/**
 * Interface for Excel context
 */
export interface ExcelContext {
  activeSheet: string;
  usedRange: {
    address: string;
    values: any[][];
  };
  selectedRange?: {
    address: string;
    values: any[][];
  };
  sheets: string[];
}

/**
 * Interface for command result
 */
export interface CommandResult {
  success: boolean;
  message: string;
  stepResults: {
    stepId: string;
    success: boolean;
    message: string;
    data?: any;
  }[];
  statusUpdates: string[];
}

/**
 * Interface for extracted document data
 */
export interface ExtractedData {
  tables: Array<{
    headers: string[];
    rows: any[][];
    title?: string;
    description?: string;
  }>;
  keyValuePairs: Record<string, any>;
  text: string;
}

// Derive base URL for knowledge base operations by stripping the unified search suffix
const API_BASE_URL = config.knowledgeBaseApiUrl.replace("/search/unified", "");

/**
 * Sends a request to the Cori API
 * @param endpoint The API endpoint
 * @param data The data to send
 * @returns The API response
 */
async function sendRequest(endpoint: string, data: any): Promise<any> {
  try {
    const response = await fetch(`${API_BASE_URL}${endpoint}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || 'An error occurred');
    }

    return await response.json();
  } catch (error) {
    console.error(`Error sending request to ${endpoint}:`, error);
    throw error;
  }
}

/**
 * Processes a user command
 * @param message User message
 * @param context Current Excel context
 * @returns Command result and status updates
 */
export async function processCommand(message: string, context: ExcelContext): Promise<CommandResult> {
  return sendRequest('/command', { message, context });
}

/**
 * Extracts data from a document
 * @param file Document file
 * @returns Extracted data and preview
 */
export async function extractDocumentData(file: File): Promise<{
  extractedData: ExtractedData;
  preview: string;
}> {
  const formData = new FormData();
  formData.append('document', file);

  try {
    const response = await fetch(`${API_BASE_URL}/extract-document`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || 'An error occurred');
    }

    return await response.json();
  } catch (error) {
    console.error('Error extracting document data:', error);
    throw error;
  }
}

/**
 * Gets context-aware suggestions for a command
 * @param partialCommand Partial command
 * @param context Current Excel context
 * @returns Suggestions
 */
export async function getSuggestions(partialCommand: string, context: ExcelContext): Promise<string[]> {
  const response = await sendRequest('/suggestions', { partialCommand, context });
  return response.suggestions;
}

/**
 * Checks if the Cori server is running
 * @returns True if the server is running, false otherwise
 */
export async function checkServerHealth(): Promise<boolean> {
  try {
    const response = await fetch(`${API_BASE_URL}/health`);
    return response.ok;
  } catch (error) {
    console.error('Knowledge-base health check failed:', error);
    return false;
  }
}

/**
 * Gets the current Excel context
 * @returns Excel context
 */
export async function getCurrentExcelContext(): Promise<ExcelContext> {
  return Excel.run(async (context) => {
    // Get the active worksheet
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    
    // Get all worksheets
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    
    // Get the used range
    const usedRange = sheet.getUsedRange();
    usedRange.load(['address', 'values']);
    
    // Get the selected range if any
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load(['address', 'values']);
    
    await context.sync();
    
    // Build the context object
    const excelContext: ExcelContext = {
      activeSheet: sheet.name,
      usedRange: {
        address: usedRange.address,
        values: usedRange.values,
      },
      sheets: sheets.items.map(s => s.name),
    };
    
    // Add selected range if it exists
    try {
      excelContext.selectedRange = {
        address: selectedRange.address,
        values: selectedRange.values,
      };
    } catch (error) {
      // No selection, ignore
    }
    
    return excelContext;
  });
}

/**
 * Executes a client action returned from the server
 * @param action Client action to execute
 * @returns Result of the execution
 */
export async function executeClientAction(action: any): Promise<any> {
  if (!action || !action.type) {
    throw new Error('Invalid client action');
  }
  
  return Excel.run(async (context) => {
    let result;
    
    switch (action.type) {
      case 'populateRange':
        result = await populateRange(context, action.range, action.data);
        break;
      case 'formatRange':
        result = await formatRange(context, action.range, action.format);
        break;
      case 'createChart':
        result = await createChart(context, action.dataRange, action.chartType, action.position);
        break;
      default:
        throw new Error(`Unknown client action type: ${action.type}`);
    }
    
    await context.sync();
    return result;
  });
}

/**
 * Populates a range with data
 */
async function populateRange(context: Excel.RequestContext, rangeAddress: string, data: any[][]): Promise<any> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(rangeAddress);
  range.values = data;
  return { message: `Range ${rangeAddress} populated with data` };
}

/**
 * Formats a range
 */
async function formatRange(context: Excel.RequestContext, rangeAddress: string, format: any): Promise<any> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(rangeAddress);
  
  if (format.bold !== undefined) {
    range.format.font.bold = format.bold;
  }
  
  if (format.italic !== undefined) {
    range.format.font.italic = format.italic;
  }
  
  if (format.color) {
    range.format.font.color = format.color;
  }
  
  if (format.fill) {
    range.format.fill.color = format.fill;
  }
  
  if (format.horizontalAlignment) {
    range.format.horizontalAlignment = format.horizontalAlignment;
  }
  
  if (format.verticalAlignment) {
    range.format.verticalAlignment = format.verticalAlignment;
  }
  
  return { message: `Range ${rangeAddress} formatted` };
}

/**
 * Creates a chart
 */
async function createChart(
  context: Excel.RequestContext, 
  dataRange: string, 
  chartType: Excel.ChartType, 
  position: { x: number, y: number, width: number, height: number }
): Promise<any> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(dataRange);
  
  const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
  
  chart.setPosition(String(position.x), String(position.y));
  chart.width = position.width;
  chart.height = position.height;
  
  return { message: `Chart created from data range ${dataRange}` };
}
