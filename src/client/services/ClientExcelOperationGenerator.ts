// src/client/services/ClientExcelOperationGenerator.ts
// Generates Excel operations using the Anthropic API

import { v4 as uuidv4 } from 'uuid';
import { ClientAnthropicService, ModelType } from './ClientAnthropicService';
import { ExcelCommandPlan, ExcelOperation } from '../models/ExcelOperationModels';

/**
 * Service that generates Excel operations using the Anthropic API
 */
export class ClientExcelOperationGenerator {
  private anthropic: ClientAnthropicService;
  private debugMode: boolean;

  constructor(params: {
    anthropic: ClientAnthropicService;
    debugMode?: boolean;
  }) {
    this.anthropic = params.anthropic;
    this.debugMode = params.debugMode || false;
  }

  /**
   * Generate Excel operations from a user query and workbook context
   * @param query The user query
   * @param workbookContext The workbook context
   * @returns A command plan with operations
   */
  public async generateOperations(
    query: string,
    workbookContext: string,
    chatHistory: Array<{ role: string; content: string }>
  ): Promise<ExcelCommandPlan> {
    try {
      // Create a system prompt for generating Excel operations
      const systemPrompt = this.buildSystemPrompt();
      
      // Use the standard model for generating operations
      const modelToUse = this.anthropic.getModel(ModelType.Standard);
      
      if (this.debugMode) {
        console.log('Generating Excel operations:', {
          model: modelToUse,
          query: query.substring(0, 50) + (query.length > 50 ? '...' : '')
        });
      }
      
      // Filter the chat history to only include the last 5 messages
      // Format the chat history for context, filtering out system messages
      const filteredChatHistory = chatHistory.slice(-5).filter(msg => msg.role !== 'system');
      // select only messages that have role user or assistant
      const messageHistory = filteredChatHistory.filter(msg => msg.role === 'user' || msg.role === 'assistant');
      
      const userPrompt = `User query: ${query}. Here is the workbook context to reference while generating operations: ${workbookContext}`;
      // Convert messageHistory to Anthropic message format
      
      const anthropicMessages = messageHistory.map(msg => ({
        role: msg.role as 'user' | 'assistant',
        content: msg.content
      }));
      anthropicMessages.push({ role: 'user' as const, content: userPrompt });
      
      // Call the API to generate operations
      const response = await this.anthropic.getClient().messages.create({
        model: modelToUse,
        system: systemPrompt,
        messages: anthropicMessages,
        max_tokens: 4000,
        temperature: 0.2 // Low temperature for more deterministic results
      });
      
      // Extract the response content
      let responseContent = response.content?.[0]?.type === 'text' 
        ? response.content[0].text 
        : '{"description":"Error generating operations","operations":[]}';
      
      try {
        // Use the extractJsonFromMarkdown utility to extract JSON from the response
        // This handles cases where the LLM includes markdown formatting or additional text
        responseContent = this.anthropic.extractJsonFromMarkdown(responseContent);
        
        if (this.debugMode) {
          console.log('Extracted JSON from response:', responseContent.substring(0, 100) + '...');
        }
        
        // Parse the extracted JSON response
        const plan = JSON.parse(responseContent) as ExcelCommandPlan;
        
        if (this.debugMode) {
          console.log('Generated Excel operations:', plan);
        }
        
        // Validate the operations
        this.validateOperations(plan.operations);
        
        return {
          description: plan.description || 'Excel operations',
          operations: plan.operations || []
        };
      } catch (parseError) {
        console.error('Failed to parse operations JSON:', parseError);
        // Return an empty plan if parsing fails
        return {
          description: 'Error parsing operations',
          operations: []
        };
      }
    } catch (error: any) {
      console.error('Error generating Excel operations:', error);
      return {
        description: 'Error generating operations',
        operations: []
      };
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
   * Build the system prompt for generating Excel operations
   * @returns The system prompt
   */
  private buildSystemPrompt(): string {
    return `You are an expert Excel assistant that generates operations for Excel workbooks. Your task is to analyze user queries and generate a list of Excel operations to fulfill their requests.

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
}

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
- rename_sheet: Rename a worksheet
- copy_range: Copy a range to another location
- merge_cells: Merge cells
- unmerge_cells: Unmerge cells
- conditional_format: Add conditional formatting
- add_comment: Add a comment to a cell

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

11. rename_sheet:
{
  "op": "rename_sheet",
  "oldName": string,      // Current sheet name
  "newName": string       // New sheet name
}

12. copy_range:
{
  "op": "copy_range",
  "source": string,       // Source range (e.g. "Sheet1!A1:D10")
  "destination": string   // Destination cell (e.g. "Sheet2!A1")
}

13. merge_cells:
{
  "op": "merge_cells",
  "range": string         // Range to merge (e.g. "Sheet1!A1:D1")
}

14. unmerge_cells:
{
  "op": "unmerge_cells",
  "range": string         // Range to unmerge (e.g. "Sheet1!A1:D1")
}

15. conditional_format:
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

16. add_comment:
{
  "op": "add_comment",
  "target": string,       // Cell reference (e.g. "Sheet1!A1")
  "text": string          // Comment text
}

EXAMPLES:

Example 1 - Set values and add a formula:
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

Example 2 - Create a chart:
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

Example 3 - Format cells:
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

Example 4 - Create a table and sort:
User: "Create a table from data in A1:D10 and sort it by column B in descending order"
{
  "description": "Create and sort table",
  "operations": [
    {
      "op": "create_table",
      "range": "Sheet1!A1:D10",
      "hasHeaders": true
    },
    {
      "op": "sort_range",
      "range": "Sheet1!A1:D10",
      "sortBy": "B",
      "sortDirection": "descending",
      "hasHeaders": true
    }
  ]
}

Example 5 - Multiple operations:
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
1. Always use the exact operation types listed above
2. Include all required fields for each operation type
3. Make sure cell references include the sheet name (e.g. "Sheet1!A1")
4. Generate operations in the correct order for execution
5. Only include fields that are relevant to the operation
6. Use the most appropriate operation types for the task
7. Be precise with ranges and cell references`;
  }
}
