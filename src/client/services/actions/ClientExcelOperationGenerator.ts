// src/client/services/ClientExcelOperationGenerator.ts
// Generates Excel operations using the Anthropic API

import { v4 as uuidv4 } from 'uuid';
import { Attachment, ClientAnthropicService, ModelType } from '../llm/ClientAnthropicService';
import { OpenAIClientService } from '../llm/OpenAIClientService';
import { ExcelCommandPlan, ExcelOperation } from '../../models/ExcelOperationModels';
import { multimodalAnalysisService } from '../document-understanding/MultimodalAnalysisService';
import { FormattingProtocol } from '../document-understanding/FormattingModels';

/**
 * Service that generates Excel operations using the Anthropic API
 */
export class ClientExcelOperationGenerator {
  private anthropic: ClientAnthropicService;
  private openai?: OpenAIClientService;
  private debugMode: boolean;
  private apiKey: string;
  private useOpenAIFallback: boolean;
  private formattingProtocol: FormattingProtocol | null = null;

  constructor(params: {
    anthropic: ClientAnthropicService;
    openai?: OpenAIClientService;
    debugMode?: boolean;
    useOpenAIFallback?: boolean;
  }) {
    this.anthropic = params.anthropic;
    this.openai = params.openai;
    this.debugMode = params.debugMode || false;
    this.useOpenAIFallback = params.useOpenAIFallback !== undefined ? params.useOpenAIFallback : true;
    // get the api key from the .env file
    this.apiKey = process.env.ANTHROPIC_API_KEY;
    
    // Try to get the formatting protocol from the multimodal analysis service
    this.formattingProtocol = multimodalAnalysisService.getCachedFormattingProtocol();
  }

  /**
   * Generate Excel operations with multimodal content
   * @param query The user query
   * @param workbookContext The workbook context
   * @param chatHistory The chat history
   * @param attachments The attachments
   * @param retryCount The retry count
   * @param maxRetries The maximum number of retries
   * @returns The generated operations
   */
  public async generateOperationsWithMultimodal(
    query: string,
    workbookContext: string,
    chatHistory: Array<{ role: string; content: string }>,
    attachments?: Attachment[],
    retryCount: number = 0,
    maxRetries: number = 2
  ): Promise<ExcelCommandPlan> {
      // Parse the workbook context to extract formatting protocol if available
      let formattingProtocol = null;
      try {
        const parsedContext = JSON.parse(workbookContext);
        if (parsedContext.formattingProtocol) {
          formattingProtocol = parsedContext.formattingProtocol;
          console.log('Formatting protocol extracted from workbook context:', formattingProtocol.length);
        }
      } catch (parseError) {
        console.warn('Error parsing workbook context to extract formatting protocol:', parseError);
      }
      
      // Create a system prompt for generating Excel operations
      const systemPrompt = this.buildSystemPrompt(formattingProtocol);
      // Filter the chat history to only include the last 1 messages
      const filteredChatHistory = chatHistory.slice(-1).filter(msg => msg.role !== 'system');
      const messageHistory = filteredChatHistory.filter(msg => msg.role === 'user' || msg.role === 'assistant');
      // Create the final user message with multimodal content
      const userPrompt = `User query: ${query}. Here is the workbook context to reference while generating operations: ${workbookContext}`;
      
      // Prepare the content array for the final user message
      let finalUserContent: any[] = [{ type: 'text', text: userPrompt }];
       
      // Add attachments if they exist
      if (attachments && attachments.length > 0) {
        for (const attachment of attachments) {
          if (attachment.type === 'image') {
            finalUserContent.push({
              type: 'image',
              source: {
                type: 'base64',
                media_type: attachment.mimeType,
                data: attachment.content
              }
            });
          } else if (attachment.type === 'pdf') {
            finalUserContent.push({
              type: 'image', // Anthropic handles PDFs as images with the right MIME type
              source: {
                type: 'base64',
                media_type: attachment.mimeType,
                data: attachment.content
              }
            });
          }
        }
      }
      const anthropicForce = true;
      if (anthropicForce) {
        // try first with anthropic
        try {
          const operations = await this.generateOperationsWithAnthropic(finalUserContent, messageHistory, systemPrompt, retryCount, maxRetries);
          return operations;
        } catch (error) {
          console.error('Error generating operations with Anthropic:', error);
          return {
            description: 'Error generating operations',
            operations: []
          };
        }
      }
       // try first with openai
      try {        
        const operations = await this.generateOperationsWithOpenAI(query, systemPrompt,workbookContext, chatHistory, attachments);
        return operations;
      } catch (error) {
        console.error('Error generating operations with OpenAI:', error);
        // try with anthropic
        if (retryCount < maxRetries) {
          retryCount++;
          try {
            const anthropicResponse = await this.generateOperationsWithAnthropic(finalUserContent, messageHistory, systemPrompt, retryCount, maxRetries);
            return anthropicResponse;
          } catch (error) {
            console.error('Error generating operations with Anthropic:', error);
            return {
              description: 'Error generating operations',
              operations: []
            };
          }
        }        
      }
      return {
        description: 'Error generating operations',
        operations: []
      };
    }


  /**
   * Generate Excel operations using OpenAI
   * @param query The user query
   * @param workbookContext The workbook context
   * @param chatHistory The chat history
   * @param attachments The attachments
   * @param retryCount The retry count
   * @param maxRetries The maximum number of retries
   * @returns The generated operations
   */

  public async generateOperationsWithOpenAI(query: string, systemPrompt: string, workbookContext: string, chatHistory: any[], attachments: Attachment[] | null)  
  {
    try {
      console.log('Attempting to generate operations with OpenAI Responses API...');
      if (this.openai && typeof (this.openai as any).generateExcelOperationsWithResponses === 'function') {
        // Use the new Responses API method if available
        return await this.openai.generateExcelOperationsWithResponses(query, systemPrompt, workbookContext, chatHistory, attachments);
      } else {
        // Fall back to the regular method if needed
        console.log('Falling back to standard OpenAI API...');
        return await this.openai.generateExcelOperations(query, systemPrompt, workbookContext, chatHistory, attachments);
      }
    } catch (openaiError) {
      console.error('Operation generation with OpenAI failed:', openaiError);
      throw openaiError;
    }
  }

  /**
   * Generate Excel operations using Anthropic
   * @param finalUserContent The final user content
   * @param messageHistory The message history
   * @param systemPrompt The system prompt
   * @param retryCount The retry count
   * @param maxRetries The maximum number of retries
   * @returns The generated operations
   */

  public async generateOperationsWithAnthropic(finalUserContent: any[],  messageHistory: any[], systemPrompt: string, retryCount: number = 0, maxRetries: number = 0)  
  {
      const modelToUse = this.anthropic.getModel(ModelType.Standard);    
      
      // Convert messageHistory to Anthropic message format
      const anthropicMessages = messageHistory.map(msg => ({
        role: msg.role as 'user' | 'assistant',
        content: [{ type: 'text', text: msg.content }]
      }));
      
      // Add the final user message to the messages array
      anthropicMessages.push({
        role: 'user',
        content: finalUserContent
      });
      
      // Use the Anthropic client from ClientAnthropicService
      // This client is already initialized with dangerouslyAllowBrowser: true      
      try{
        const anthropicClient = this.anthropic.getClient();
        const responseData = await anthropicClient.messages.create({
          model: modelToUse,
          system: systemPrompt,
          messages: anthropicMessages as any, // Type assertion to resolve SDK type issue
          max_tokens: 4000,
          temperature: 0.2
        });
        
        // Extract the response content
        let responseContent = responseData.content?.[0]?.type === 'text' 
          ? responseData.content[0].text 
          : '{"description":"Error generating operations","operations":[]}';
        
        try {
          // Use the extractJsonFromMarkdown utility to extract JSON from the response
          responseContent = this.anthropic.extractJsonFromMarkdown(responseContent);
          
          // Parse the extracted JSON response
          const plan = JSON.parse(responseContent) as ExcelCommandPlan;
          
          return {
            description: plan.description || 'Excel operations',
            operations: plan.operations || []
          };
        } catch (parseError) {
          console.error(`Failed to parse operations JSON (attempt ${retryCount + 1}/${maxRetries + 1}):`, parseError);
          throw parseError;
        }
      }
      catch(error){
        console.error('Error generating operations with Anthropic:', error);
        throw error;
      }
    }

  /**
   * Validate the operations to ensure they are well-formed
   * @param operations The operations to validate
   */
  private validateOperations(operations: ExcelOperation[]): void {
    if (!operations || !Array.isArray(operations)) {
      throw new Error('Operations must be an array');
    }
    
    for (const operation of operations) {
      if (!operation.op) {
        throw new Error('Operation missing "op" field');
      }
      
      // Additional validation could be added here based on operation type
    }
  }
  
  /**
   * Build the system prompt for the Excel operation generator
   * @param formattingProtocol The formatting protocol (optional)
   * @param isRetry Whether this is a retry attempt after a failed parsing
   * @returns The system prompt
   */
  private buildSystemPrompt(formattingProtocol: FormattingProtocol | null = null, isRetry: boolean = false): string {
    let basePrompt = `You are an expert Excel assistant that generates operations for Excel workbooks. Your task is to analyze user queries and generate a list of Excel operations to fulfill their requests.

CRITICAL INSTRUCTION: ONLY generate operations that the user EXPLICITLY asks for. DO NOT add any additional operations that the user did not request. If the user asks to "add a new tab", ONLY create a new worksheet and DO NOT add any data, charts, or formatting to it unless specifically requested.`;
    
    // Add more explicit instructions for retry attempts
    if (isRetry) {
      basePrompt = `${basePrompt}

CRITICAL JSON FORMATTING INSTRUCTION: Your previous response contained invalid JSON that could not be parsed. You MUST respond with ONLY a valid JSON object in a code block. Do not include any explanations, notes, or text outside the JSON code block. The JSON must exactly follow the schema provided below with no extra or missing fields. Ensure all quotes, brackets, and commas are properly placed.`;
    }
    
    // Update the formatting protocol in the system prompt if available
    if (formattingProtocol) {
      console.log('Using formatting protocol in system prompt');
      basePrompt += `

FORMATTING PROTOCOL:
${JSON.stringify(formattingProtocol, null, 2)}`;
      if (formattingProtocol.colorCoding) {
        basePrompt += `
COLOR CODING CONVENTIONS:
`;
        
        if (formattingProtocol.colorCoding.inputs) {
          basePrompt += `- Input cells: ${formattingProtocol.colorCoding.inputs}
`;
        }
        if (formattingProtocol.colorCoding.calculations) {
          basePrompt += `- Calculation cells: ${formattingProtocol.colorCoding.calculations}
`;
        }
        if (formattingProtocol.colorCoding.hardcodedValues) {
          basePrompt += `- Hardcoded values: ${formattingProtocol.colorCoding.hardcodedValues}
`;
        }
        if (formattingProtocol.colorCoding.linkedValues) {
          basePrompt += `- Linked values: ${formattingProtocol.colorCoding.linkedValues}
`;
        }
        if (formattingProtocol.colorCoding.headers) {
          basePrompt += `- Headers: ${formattingProtocol.colorCoding.headers}
`;
        }
        if (formattingProtocol.colorCoding.totals) {
          basePrompt += `- Totals: ${formattingProtocol.colorCoding.totals}
`;
        }
      }
      
      // Add number formatting instructions
      if (formattingProtocol.numberFormatting) {
        basePrompt += `
NUMBER FORMATTING CONVENTIONS:
`;
        
        if (formattingProtocol.numberFormatting.currency) {
          basePrompt += `- Currency: ${formattingProtocol.numberFormatting.currency}
`;
        }
        if (formattingProtocol.numberFormatting.percentage) {
          basePrompt += `- Percentage: ${formattingProtocol.numberFormatting.percentage}
`;
        }
        if (formattingProtocol.numberFormatting.date) {
          basePrompt += `- Date: ${formattingProtocol.numberFormatting.date}
`;
        }
        if (formattingProtocol.numberFormatting.general) {
          basePrompt += `- General: ${formattingProtocol.numberFormatting.general}
`;
        }
      }
      
      // Add border styling instructions
      if (formattingProtocol.borderStyles) {
        basePrompt += `
BORDER STYLING CONVENTIONS:
`;
        
        if (formattingProtocol.borderStyles.tables) {
          basePrompt += `- Tables: ${formattingProtocol.borderStyles.tables}
`;
        }
        if (formattingProtocol.borderStyles.sections) {
          basePrompt += `- Sections: ${formattingProtocol.borderStyles.sections}
`;
        }
        if (formattingProtocol.borderStyles.totals) {
          basePrompt += `- Totals: ${formattingProtocol.borderStyles.totals}
`;
        }
      }
      
      basePrompt += `
When creating new elements in the workbook, ensure they follow these formatting conventions for consistency.`;
    }
    
    basePrompt += `

OUTPUT FORMAT:
You must respond with a JSON object that follows this schema:
{
  "description": string,  // A brief description of what these operations will do
  "operations": [         // Array of operations to execute
    {
      "op": string,       // Operation type (see allowed values below)
      ...                 // Additional fields specific to the operation type
    }
  ]
}`;
    
    return basePrompt + `

ALLOWED OPERATION TYPES:
- set_value: Set a value in a cell
- add_formula: Add a formula to a cell
- create_chart: Create a chart
- format_range: Format a range of cells
- clear_range: Clear a range of cells
- create_table: Create a table
- sort_range: Sort a range
- filter_range: Filter a range
- create_sheet: Create a new worksheet
- delete_sheet: Delete a worksheet
- copy_range: Copy a range to another location
- merge_cells: Merge cells
- unmerge_cells: Unmerge cells
- conditional_format: Add conditional formatting
- add_comment: Add a comment to a cell
- set_freeze_panes: Freeze rows or columns
- set_active_sheet: Set the active worksheet
- set_print_settings: Set print settings
- set_page_setup: Set page setup
- export_to_pdf: Export worksheet to PDF
- set_worksheet_settings: Set worksheet settings
- format_chart: Format a chart

OPERATION SCHEMAS:

1. set_value:
{
  "op": "set_value",
  "target": string,       // Cell reference (e.g. "Sheet1!A1")
  "value": any            // Value to set (string, number, boolean)
}

2. add_formula:
{
  "op": "add_formula",
  "target": string,       // Cell reference (e.g. "Sheet1!A1")
  "formula": string       // Formula to add (e.g. "=SUM(B1:B10)")
}

3. create_chart:
{
  "op": "create_chart",
  "range": string,        // Range for chart data (e.g. "Sheet1!A1:D10")
  "type": string,         // Chart type (e.g. "columnClustered", "line", "pie")
  "title": string,        // Optional chart title
  "position": string      // Optional position (e.g. "Sheet1!F1")
}

4. format_range:
{
  "op": "format_range",
  "range": string,        // Range to format (e.g. "Sheet1!A1:D10")
  "style": string,        // Optional number format (e.g. "Currency", "Percentage")
  "bold": boolean,        // Optional bold formatting
  "italic": boolean,      // Optional italic formatting
  "fontColor": string,    // Optional font color
  "fillColor": string,    // Optional fill color
  "fontSize": number,     // Optional font size
  "horizontalAlignment": string, // Optional alignment (e.g. "left", "center", "right")
  "verticalAlignment": string    // Optional alignment (e.g. "top", "center", "bottom")
}

5. clear_range:
{
  "op": "clear_range",
  "range": string,        // Range to clear (e.g. "Sheet1!A1:D10")
  "clearType": string     // Optional clear type (e.g. "all", "formats", "contents")
}

6. create_table:
{
  "op": "create_table",
  "range": string,        // Range for table (e.g. "Sheet1!A1:D10")
  "hasHeaders": boolean,  // Optional whether first row contains headers
  "styleName": string     // Optional table style name
}

7. sort_range:
{
  "op": "sort_range",
  "range": string,        // Range to sort (e.g. "Sheet1!A1:D10")
  "sortBy": string,       // Column to sort by (e.g. "A", "B", "C")
  "sortDirection": string, // Sort direction (e.g. "ascending", "descending")
  "hasHeaders": boolean   // Optional whether first row contains headers
}

8. filter_range:
{
  "op": "filter_range",
  "range": string,        // Range to filter (e.g. "Sheet1!A1:D10")
  "column": string,       // Column to filter (e.g. "A", "B", "C")
  "criteria": string      // Filter criteria (e.g. ">0", "=Red", "<>0")
}

9. create_sheet:
{
  "op": "create_sheet",
  "name": string,         // Name for the new sheet
  "position": number      // Optional position (0-based index)
}

10. delete_sheet:
{
  "op": "delete_sheet",
  "name": string          // Name of the sheet to delete
}

11. copy_range:
{
  "op": "copy_range",
  "source": string,       // Source range (e.g. "Sheet1!A1:D10")
  "destination": string   // Destination cell (e.g. "Sheet2!A1")
}

12. merge_cells:
{
  "op": "merge_cells",
  "range": string         // Range to merge (e.g. "Sheet1!A1:D1")
}

13. unmerge_cells:
{
  "op": "unmerge_cells",
  "range": string         // Range to unmerge (e.g. "Sheet1!A1:D1")
}

14. conditional_format:
{
  "op": "conditional_format",
  "range": string,        // Range to format (e.g. "Sheet1!A1:D10")
  "type": string,         // Format type (e.g. "dataBar", "colorScale", "iconSet", "topBottom", "custom")
  "criteria": string,     // Optional criteria for custom formats
  "format": {             // Optional format settings for custom formats
    "fontColor": string,
    "fillColor": string,
    "bold": boolean,
    "italic": boolean
  }
}

15. add_comment:
{
  "op": "add_comment",
  "target": string,       // Cell reference (e.g. "Sheet1!A1")
  "text": string          // Comment text
}

16. set_freeze_panes:
{
  "op": "set_freeze_panes",
  "sheet": string,         // Sheet name
  "address": string        // Cell address to freeze at (e.g. "B3")
  "freeze": boolean        // Whether to freeze panes. True if the user wants to freeze panes and false if they want to unfreeze.
}

17. set_print_settings:
{
  "op": "set_print_settings",
  "sheet": string,         // Sheet name
  blackAndWhite: boolean,  // Whether to print in black and white
  draftMode: boolean,      // Whether to print in draft mode
  firstPageNumber: number, // First page number
  headings: boolean,       // Whether to display row/column headings when printing
  orientation: string,     // "portrait" or "landscape"
  printAreas: string[],    // Ranges to set as print areas (e.g. ["A1:H20", "A20:H40"])
  printComments: string,   // "none", "at_end", "as_displayed"
  headerMargin: number,    // Header margin in inches
  footerMargin: number,    // Footer margin in inches
  leftMargin: number,      // Left margin in inches
  rightMargin: number,     // Right margin in inches
  topMargin: number,       // Top margin in inches
  bottomMargin: number,    // Bottom margin in inches
  printErrors: string,     // "blank", "dash", "displayed", "na"
  headerRows: number,      // Number of header rows
  footerRows: number,      // Number of footer rows 
  printTitles: string[],   // Ranges to set as print titles (e.g. ["A1:H1", "A1:H1"])
  printGridlines: boolean, // Whether to display gridlines when printing
}

18. set_page_setup:
{
  "op": "set_page_setup",
  "sheet": string,         // Sheet name
  "pageLayoutView": string,    // "print", "normal", "pageBreakPreview"
  "zoom": number,          // Zoom percentage
  "gridlines": boolean,    // Whether to display gridlines
  "headers": boolean,      // Whether to display row and column headers
  "showFormulas": boolean, // Whether to display formulas instead of values
  "showHeadings": boolean  // Whether to display row and column headings
}

19. export_to_pdf:
{
  "op": "export_to_pdf",
  "sheet": string,         // Sheet name to export
  "fileName": string,      // Optional name for the PDF file (without extension)
  "quality": string,       // Optional PDF quality: "standard" or "minimal"
  "includeComments": boolean, // Optional: whether to include comments
  "printArea": string,     // Optional print area to export (e.g., "A1:H20")
  "orientation": string,   // Optional page orientation: "portrait" or "landscape"
  "fitToPage": boolean,    // Optional: whether to fit content to page
  "margins": {             // Optional page margins in points
    "top": number,
    "right": number,
    "bottom": number,
    "left": number
  }
}

20. set_worksheet_settings:
{
  "op": "set_worksheet_settings",
  "sheet": string,         // Sheet name
  "pageLayoutView": string,    // "print", "normal", "pageBreakPreview"
  "zoom": number,          // Zoom percentage
  "gridlines": boolean,    // Whether to display gridlines
  "headers": boolean,      // Whether to display row and column headers
  "showFormulas": boolean, // Whether to display formulas instead of values
  "showHeadings": boolean  // Whether to display row and column headings
  "position": int          // index of the sheet in the whole workbook. 0 based.
  "enableCalculation": boolean // Whether to enable calculation
  "visibility": boolean    // Whether to make the sheet visible
}

21. format_chart:
{
  "op": "format_chart",
  "sheet": string,         // Sheet name
  "chart": string,         // Chart name
  "title": string,         // Chart title
  "type": string,          // Chart type
  "dataSource": string,    // Chart data source cell range address
  "legend": boolean,        // Chart legend
  "axis": boolean,          // Chart axis
  "series": string,        // Chart series
  "dataLabels": boolean,    // Chart data labels
  
  // Chart Dimension properties
  "width": number,         // Chart width
  "height": number,        // Chart height
  
  // Chart position properties
  "left": number,          // Chart left position
  "top": number,           // Chart top position

    // Chart format properties
  "fillColor": string, // Chart fill color
  "borderVisible": boolean, // Chart border visibility
  "borderColor": string, // Chart border color
  "borderWidth": number, // Chart border width
  "borderStyle": string, // Chart border style
  "borderDashStyle": string, // Chart border dash style
  
  // Chart title properties
  "titleVisible": boolean, // Whether title is visible
  "titleFontName": string, // Title font name
  "titleFontSize": number, // Title font size
  "titleFontStyle": string, // Title font style
  "titleFontBold": boolean, // Title font bold
  "titleFontItalic": boolean, // Title font italic
  "titleFontColor": string, // Title font color
  "titleFormat": string, // Title format

  // Legend properties
  "legendVisible": boolean, // Whether legend is visible
  "legendFontName": string, // Legend font name
  "legendFontSize": number, // Legend font size
  "legendFontStyle": string, // Legend font style
  "legendFontBold": boolean, // Legend font bold
  "legendFontItalic": boolean, // Legend font italic
  "legendFontColor": string, // Legend font color
  "legendFormat": string, // Legend format

  // Chart axis properties
  "axisVisible": boolean, // Whether axis is visible
  "axisFontName": string, // Axis font name
  "axisFontSize": number, // Axis font size
  "axisFontStyle": string, // Axis font style
  "axisFontBold": boolean, // Axis font bold
  "axisFontItalic": boolean, // Axis font italic
  "axisFontColor": string, // Axis font color
  "axisFormat": string, // Axis format

  // Chart series properties
  "seriesVisible": boolean, // Whether series is visible
  "seriesFontName": string, // Series font name
  "seriesFontSize": number, // Series font size
  "seriesFontStyle": string, // Series font style
  "seriesFontBold": boolean, // Series font bold
  "seriesFontItalic": boolean, // Series font italic
  "seriesFontColor": string, // Series font color
  "seriesFormat": string, // Series format

  // Chart data labels properties
  "dataLabelsVisible": boolean, // Whether data labels are visible
  "dataLabelsFontName": string, // Data labels font name
  "dataLabelsFontSize": number, // Data labels font size
  "dataLabelsFontStyle": string, // Data labels font style
  "dataLabelsFontBold": boolean, // Data labels font bold
  "dataLabelsFontItalic": boolean, // Data labels font italic
  "dataLabelsFontColor": string, // Data labels font color
  "dataLabelsFormat": string, // Data labels format

}

22. set_calculation_options:
{
  "op": "set_calculation_options",
  "calculationMode": string, // Optional calculation mode (e.g. "auto", "manual")
  "iterative": boolean,      // Optional whether to enable iterative calculation
  "maxIterations": number,   // Optional maximum number of iterations for iterative calculation
  "maxChange": number,       // Optional maximum change for iterative calculation
  "calculate": boolean,      // Optional whether to calculate the workbook
  "calculationType": string  // Optional calculation type (e.g. "full", "full_recalculate", "recalculate")
}

23. recalculate_ranges:
  {
    "op": "recalculate_ranges",
    "recalculateAll": boolean, // Optional whether to recalculate all sheets
    "sheets": string[] // List of sheet names to recalculate
    "ranges": string[] // List of cell range addresses to recalculate
  }

EXAMPLES:

Example 1 - Create a new worksheet (minimal operation):
User: "Add a new tab called Sales"
{
  "description": "Create new Sales worksheet",
  "operations": [
    {
      "op": "create_sheet",
      "name": "Sales"
    }
  ]
}

Example 2 - Set values and add a formula (only what's requested):
User: "Put 10 in cell A1, 20 in cell A2, and calculate the sum in A3"
{
  "description": "Set values and calculate sum",
  "operations": [
    {
      "op": "set_value",
      "target": "Sheet1!A1",
      "value": 10
    },
    {
      "op": "set_value",
      "target": "Sheet1!A2",
      "value": 20
    },
    {
      "op": "add_formula",
      "target": "Sheet1!A3",
      "formula": "=SUM(A1:A2)"
    }
  ]
}

Example 3 - Create a chart (only what's requested):
User: "Create a column chart for sales data in range A1:B10 with the title 'Sales Report'"
{
  "description": "Create sales chart",
  "operations": [
    {
      "op": "create_chart",
      "range": "Sheet1!A1:B10",
      "type": "columnClustered",
      "title": "Sales Report",
      "position": "Sheet1!D1"
    }
  ]
}

Example 4 - Format cells (only what's requested):
User: "Format cells B2:B10 as currency and make them bold"
{
  "description": "Format cells as currency and bold",
  "operations": [
    {
      "op": "format_range",
      "range": "Sheet1!B2:B10",
      "style": "Currency",
      "bold": true
    }
  ]
}

Example 5 - Freeze panes (single operation):
User: "Freeze the first row"
{
  "description": "Freeze first row",
  "operations": [
    {
      "op": "set_freeze_panes",
      "sheet": "Sheet1",
      "row": 1,
      "column": 0
    }
  ]
}

Example 5a - Freeze panes using cell address:
User: "Freeze panes at cell B3"
{
  "description": "Freeze panes at cell B3",
  "operations": [
    {
      "op": "set_freeze_panes",
      "sheet": "Sheet1",
      "address": "B3"      // Use address for cell reference instead of row/column
    }
  ]
}

Example 6 - Multiple explicitly requested operations:
User: "Create a new sheet called 'Summary', copy data from Sheet1!A1:D10 to Summary!A1, and format as currency"
{
  "description": "Create summary sheet with formatted data",
  "operations": [
    {
      "op": "create_sheet",
      "name": "Summary"
    },
    {
      "op": "copy_range",
      "source": "Sheet1!A1:D10",
      "destination": "Summary!A1"
    },
    {
      "op": "format_range",
      "range": "Summary!A1:D10",
      "style": "Currency"
    }
  ]
}

Important rules:
1. ONLY generate operations that the user EXPLICITLY requests - this is the most important rule
2. Keep the number of operations to the absolute minimum required to fulfill the user's request
3. Do not add any "helpful" operations that weren't requested
4. Do not populate new worksheets with data unless specifically requested
5. Always use the exact operation types listed above
6. Include all required fields for each operation type
7. Make sure cell references include the sheet name (e.g. "Sheet1!A1")
8. Generate operations in the correct order for execution
9. Only include fields that are relevant to the operation
10. Use the most appropriate operation types for the task
11. Be precise with ranges and cell references

REMEMBER: If the user asks for a simple operation like "add a new tab named Sales", your response should ONLY include that specific operation and nothing more. Do not add any data, formatting, or additional operations.

ANTI-PATTERNS TO AVOID:
1. DO NOT create sample data unless explicitly requested
2. DO NOT create charts or visualizations unless explicitly requested
4. DO NOT create multiple sheets when only one was requested
5. DO NOT add operations that seem "helpful" but weren't requested
6. DO NOT respond with string responses like "I'll perform [action]" or "I'll do [this operation]". Remember that you have to respond only with the JSON schema.


When in doubt, be minimalist and only do exactly what was asked.`;
  }
}
