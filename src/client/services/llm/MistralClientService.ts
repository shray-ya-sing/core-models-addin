/**
 * MistralClientService.ts
 * 
 * Client service for interacting with Mistral AI models
 */

import { Mistral } from '@mistralai/mistralai';
import { v4 as uuidv4 } from 'uuid';

/**
 * Enum for Mistral model types
 */
export enum MistralModelType {
  Light = 'light',
  Standard = 'standard',
  Advanced = 'advanced',
  Fast = 'fast'
}

/**
 * Client service for interacting with Mistral AI
 */
export class MistralClientService {
  private mistral: Mistral;
  private models: { [key in MistralModelType]: string };
  private debugMode: boolean;
  
  /**
   * Constructor
   * @param apiKey Mistral API key
   * @param debugMode Enable debug mode for verbose logging
   */
  constructor(debugMode = false) {
    this.mistral = new Mistral({
      apiKey: 'XHst9dbFgRSRoG5s4P9vawDk3Cn168tn',
    });
    
    this.models = {
      [MistralModelType.Light]: 'mistral-small-latest',
      [MistralModelType.Standard]: 'pixtral-12b-2409',
      [MistralModelType.Advanced]: 'pixtral-large-latest',
      [MistralModelType.Fast]: 'mistral-small-latest'
    };
    
    this.debugMode = debugMode;
  }
  
  /**
   * Use Mistral to select relevant sheets based on the user's query
   * @param query The user's query
   * @param availableSheets List of available sheets in the workbook
   * @param chatHistory Chat history for context
   * @returns Array of sheet names that are relevant to the query
   */
  public async selectRelevantSheets(
    query: string,
    availableSheets: Array<{name: string, summary: string}>,
    chatHistory: Array<{role: string, content: string, attachments?: any[]}>
  ): Promise<string[]> {
    try {
      // Enhanced debug logging
      if (this.debugMode) {
        console.log(
          '%c MistralClientService: Selecting sheets for query: ' + 
          `"${query}"`,
          'background: #7e22ce; color: white; font-weight: bold; padding: 2px 5px;'
        );
        
        console.log(`%c Available sheets: ${availableSheets.map(s => s.name).join(', ')}`, 'color: #7e22ce;');
      }
      
      // Format the available sheets as a list
      const sheetsDescription = availableSheets.map(sheet => 
        `- "${sheet.name}": ${sheet.summary || 'No summary available'}`
      ).join('\n');
      
      // Create a clear, structured prompt for sheet selection
      const systemPrompt = `You are an Excel assistant that helps users find relevant sheets in their workbook.
      
YOUR TASK:
1. Given a user's query about an Excel workbook and a list of available sheets
2. Determine which sheets are most likely relevant to answering their query
3. Return ONLY a JSON array of sheet names, with no other text or explanation
4. Select sheet based only on the most recent query

You should prefer to include sheets when:
- The sheet name is explicitly mentioned in the query
- The sheet contains data that would be needed to answer the query
- The sheet's purpose aligns with the query's subject matter

IF THE USER REQUESTS TO ADD A NEW SHEET OR A SHEET THAT DOES NOT EXIST, THEN SELECT ALL EXISTING SHEETS.

SPECIAL INSTRUCTIONS FOR WORKBOOK-LEVEL QUERIES:
- If the query is about the entire workbook (examples: explain the workbook, overview, how many sheets, etc.)
- Or if the query requires context from multiple sheets to answer properly
- Or if you're unsure whether the query needs one sheet or multiple sheets
THEN include ALL sheets in your response.

Here's some context about which data is typically in which sheet to help you make a decision about which sheets to include: 

Income Statement: Revenue, Expenses, Profit, EBITDA, etc.
Balance Sheet: Assets, Liabilities, Equity, etc.
Cash Flow: Operating activities, Investing activities, Financing activities, etc.

The user's workbook may have abbreviated tabs 'BS' for balance sheet, 'IS' for income statement, and so on.
Be conservative in your selection of sheets.

RESPOND WITH VALID JSON ONLY - an array of strings representing sheet names.`;
      
      // Format the chat history for context, filtering out system messages
      const filteredChatHistory = chatHistory.filter(msg => msg.role !== 'system');
      const chatHistoryContext = filteredChatHistory.length > 0 ?
        `\nCHAT HISTORY FOR CONTEXT:\n${filteredChatHistory.map(msg => `${msg.role.toUpperCase()}: ${msg.content}`).join('\n')}` :
        '';
      
      const userPrompt = `USER QUERY: "${query}"

AVAILABLE SHEETS:
${sheetsDescription}${chatHistoryContext}

Return a JSON array containing ONLY the names of sheets relevant to the query.`;

      // Create the message structure for the Mistral API
      const messages = [
        { role: 'system' as const, content: systemPrompt },
        { role: 'user' as const, content: userPrompt }
      ];
      
      // Use the Fast model as requested
      const modelToUse = this.models[MistralModelType.Fast];
      
      // Make the API call
      const response = await this.mistral.chat.complete({
        model: modelToUse,
        messages: messages,
        temperature: 0.1  // Low temperature for consistent results
      });
      
      // Extract the response content
      const messageText = response.choices[0]?.message?.content || '[]';
      const contentText = typeof messageText === 'string' ? messageText : '[]';
      
      // Extract JSON array from the response
      const jsonText = this.extractJsonFromMarkdown(contentText);
      
      try {
        const result = JSON.parse(jsonText);
        
        // Check if the result contains a 'sheets' field or is a direct array
        const sheetNames = Array.isArray(result) ? result : 
                          (result.sheets && Array.isArray(result.sheets)) ? result.sheets : 
                          [];
        
        // Log the selected sheets
        if (this.debugMode) {
          console.log(`%c Mistral selected sheets: ${sheetNames.join(', ')}`, 'color: #7e22ce;');
        }
        
        return sheetNames;
      } catch (parseError) {
        console.error('Error parsing sheet selection response:', parseError);
        console.log('Raw response:', jsonText);
        return [];
      }
    } catch (error: any) {
      console.error('Error selecting relevant sheets with Mistral:', error);
      // If there's an error, return empty array
      return [];
    }
  }
  
  /**
   * Extract JSON from markdown text
   * @param text Text that may contain JSON in a markdown code block
   * @returns Extracted JSON string
   */
  private extractJsonFromMarkdown(text: string): string {
    // Try to extract JSON from markdown code blocks
    const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)```/;
    const match = text.match(codeBlockRegex);
    
    if (match && match[1]) {
      return match[1].trim();
    }
    
    // If no code block is found, look for anything that resembles JSON
    const jsonRegex = /\[\s*".*"\s*\]|\{\s*".*"\s*:.*\}/;
    const jsonMatch = text.match(jsonRegex);
    
    if (jsonMatch) {
      return jsonMatch[0];
    }
    
    // If all else fails, return the original text
    return text;
  }
  
  /**
   * Handle API errors
   * @param error The error from the API
   * @returns A standardized error
   */
  private handleApiError(error: any): Error {
    console.error('Mistral API error:', error);
    
    if (error.status === 429) {
      return new Error('Rate limit exceeded. Please try again later.');
    }
    
    if (error.status >= 500) {
      return new Error('Mistral service is currently unavailable. Please try again later.');
    }
    
    return new Error(`Mistral API error: ${error.message || 'Unknown error'}`);
  }
}
