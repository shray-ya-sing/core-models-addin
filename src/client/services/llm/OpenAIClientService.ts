/**
 * OpenAI Client Service
 * Client-side service for interacting with the OpenAI API
 */
import { v4 as uuidv4 } from 'uuid';
import OpenAI from 'openai';
import { CommandStatus } from '../../models/CommandModels';

/**
 * Attachment type for multimodal messages
 */
export interface Attachment {
  type: 'image' | 'pdf';
  name: string;
  content: string; // base64 encoded content
  mimeType: string;
}

// Model types for different query complexities
export enum ModelType {
  Light = 'light',     // For simple queries like greetings
  Standard = 'standard', // For regular workbook queries
  Advanced = 'advanced'  // For complex operations
}

/**
 * Client-side service for interacting with the OpenAI API
 */
export class OpenAIClientService {
  private openai: OpenAI;
  private debugMode: boolean = true;
  private verboseLogging: boolean = true;
  
  // Model configuration
  private models = {
    [ModelType.Light]: 'gpt-3.5-turbo',
    [ModelType.Standard]: 'gpt-4-turbo',
    [ModelType.Advanced]: 'gpt-4-vision-preview'
  };
  
  /**
   * Creates a new instance of the OpenAIClientService
   * @param apiKey Optional OpenAI API key (will use environment variable if not provided)
   */
  constructor(apiKey?: string) {
    // Use provided API key or fall back to environment variable
    const apiKeyToUse = apiKey || process.env.OPENAI_API_KEY;
    
    if (!apiKeyToUse) {
      console.warn('No OpenAI API key provided and OPENAI_API_KEY environment variable not found.');
    }
    
    this.openai = new OpenAI({
      apiKey: apiKeyToUse,
      dangerouslyAllowBrowser: true // For client-side use
    });
  }
  
  /**
   * Create an instance of OpenAIClientService using environment variables
   * @returns A new OpenAIClientService instance
   */
  public static fromEnv(): OpenAIClientService {
    return new OpenAIClientService();
  }
  
  /**
   * Get the OpenAI client instance
   * @returns The OpenAI client
   */
  public getClient(): OpenAI {
    return this.openai;
  }
  
  /**
   * Get the appropriate model for the given complexity
   * @param type The model type to use
   * @returns The model name
   */
  public getModel(type: ModelType): string {
    return this.models[type];
  }
  
  /**
   * Handle API errors and provide a more user-friendly message
   * @param error The API error
   * @returns A user-friendly error
   */
  private handleApiError(error: any): Error {
    // Check if it's a rate limit error
    if (error.status === 429) {
      return new Error('OpenAI rate limit exceeded. Please try again in a few moments.');
    }
    
    // Check if it's an authentication error
    if (error.status === 401) {
      return new Error('Authentication error with OpenAI. Please check your API key.');
    }
    
    // Check if it's a server error
    if (error.status >= 500 && error.status < 600) {
      return new Error('OpenAI servers are experiencing issues. Please try again later.');
    }
    
    // For other errors, return a generic message with the original error
    return new Error(`Error communicating with OpenAI: ${error.message || 'Unknown error'}`);
  }
  
  /**
   * Extract JSON from a markdown string
   * @param markdownString The markdown string that may contain JSON
   * @returns The extracted JSON object, or null if none found
   */
  public extractJsonFromMarkdown(markdownString: string): any {
    try {
      // Try to parse the entire string as JSON first
      return JSON.parse(markdownString);
    } catch (e) {
      // If that fails, try to extract JSON from markdown code blocks
      const jsonRegex = /```(?:json)?\s*([\s\S]*?)\s*```/;
      const match = markdownString.match(jsonRegex);
      
      if (match && match[1]) {
        try {
          return JSON.parse(match[1]);
        } catch (e) {
          console.error('Error parsing extracted JSON:', e);
        }
      }
      
      // If no valid JSON found in code blocks, try to find any JSON-like structure
      const jsonObjectRegex = /\{[\s\S]*\}/;
      const objectMatch = markdownString.match(jsonObjectRegex);
      
      if (objectMatch && objectMatch[0]) {
        try {
          return JSON.parse(objectMatch[0]);
        } catch (e) {
          console.error('Error parsing JSON-like structure:', e);
        }
      }
      
      return null;
    }
  }
  
  /**
   * Generate a response to a workbook explanation query, using the workbook context
   * @param userQuery The user's query about the workbook
   * @param workbookContext The compressed workbook context
   * @param streamHandler Optional callback for handling streaming responses
   * @param chatHistory Optional chat history for context
   * @param attachments Optional attachments (images/pdfs)
   * @returns The generated response
   */
  public async generateWorkbookExplanation(
    userQuery: string,
    workbookContext: string,
    streamHandler?: (chunk: string) => void,
    chatHistory?: Array<{role: string, content: string, attachments?: Attachment[]}>,
    attachments?: Attachment[]
  ): Promise<any> {
    // Set a timeout for the API request (30 seconds)
    const TIMEOUT_MS = 30000;
    let timeoutId: NodeJS.Timeout;
    
    // Create a promise that rejects after the timeout
    const timeoutPromise = new Promise((_, reject) => {
      timeoutId = setTimeout(() => {
        reject(new Error('Request timed out after 30 seconds'));
      }, TIMEOUT_MS);
    });
    
    // Max retry attempts for rate limiting and other recoverable errors
    const MAX_RETRIES = 3;
    let retryCount = 0;
    let lastError: any = null;
    
    while (retryCount <= MAX_RETRIES) {
      try {
        // Variables to store response data
        let fullResponse = '';
        let messageText = '';
        let response: any;
        
        // Create a system prompt specifically for workbook explanations
        const systemPrompt = `Your name is Cori. You are an Excel assistant that helps users understand and analyze their spreadsheets. 

Analyze the provided Excel workbook context and answer the user's question in a clear, concise way.

For general workbook overview questions:
1. Provide a high-level summary of the entire workbook (1 paragraph)
2. Give a brief overview of EACH sheet (1 paragraph per sheet)
3. Explain how the sheets relate to each other
4. DO NOT provide detailed cell-by-cell analysis unless specifically asked

For sheet-specific questions:
1. Focus on the requested sheet and ignore the metadata other other sheets unless they are linked to the requested sheet
2. Explain its purpose, key data regions, and important formulas
3. Highlight any charts or tables and their significance
4. DO NOT provide detailed cell-by-cell analysis unless specifically asked

For cell-range-specific questions:
1. Focus on the requested cell range
2. Explain its purpose, key data regions, and important formulas
3. DO NOT provide information regarding other cell ranges that are not linked to this cell range
4. DO NOT provide detailed cell-by-cell analysis unless specifically asked

For single cell specific questions:
1. Focus on the requested cell
2. Explain its purpose, formula and formatting
3. DO NOT provide information regarding other cells that are not linked to this cell/ that this cell is not linked to
4. Explain the precedent values that this cell is linked to, and any dependent values that are linked to this cell

DO NOT suggest executing commands or making edits to the workbook unless explicitly requested.
Your goal is to help the user understand their existing data, not to modify it.

Format your response using proper Markdown syntax:
- Use headings (## and ###) to organize your explanation
- Use bullet points or numbered lists where appropriate
- Use **bold** or *italic* for emphasis
- Use code formatting for formulas: \`=SUM(A1:A10)\`
- Use tables for structured data where helpful

Keep your explanations CONCISE. For a full workbook overview, aim for 1-2 paragraphs per sheet maximum. If a sheet does not have any data you do not need to include it in your summary.
BE AS CONCISE AS POSSIBLE. DO NOT REPEAT CONTENT OR ADD REDUNDANT INFORMATION.
RESPOND IN AS FEW CHARACTERS AS POSSIBLE

When uncertain about any aspect, openly acknowledge limitations in your understanding rather than guessing.`;
        
        // For workbook explanations, we'll use GPT-4 which balances capability and speed
        const modelToUse = this.models[ModelType.Standard];
        
        // Log the workbook context being sent to the LLM if verbose logging is enabled
        if (this.verboseLogging) {
          try {
            // Parse the JSON to format it nicely
            const parsedContext = JSON.parse(workbookContext);
            
            // Create a collapsible console group
            console.groupCollapsed(
              '%c 📊 WORKBOOK CHUNKS SENT TO LLM 📊',
              'background: #2c3e50; color: #ecf0f1; font-size: 14px; padding: 5px 10px; border-radius: 4px;'
            );
            
            // Display basic stats
            console.log(`%c Query: "${userQuery}"`, 'color: #3498db; font-weight: bold;');
            console.log(`%c Total sheets: ${parsedContext.sheets.length}`, 'color: #2ecc71');
            console.log(`%c Active sheet: ${parsedContext.activeSheet}`, 'color: #e67e22');
            
            // Show each sheet's data in a collapsible group
            console.log('%c Included sheets:', 'color: #3498db; font-weight: bold;');
            parsedContext.sheets.forEach((sheet: any, index: number) => {
              console.groupCollapsed(`%c Sheet ${index + 1}: ${sheet.name}`, 'color: #16a085; font-weight: bold;');
              
              // Sheet summary
              if (sheet.summary) {
                console.log(`%c Summary: ${sheet.summary}`, 'color: #7f8c8d');
              }
              
              // Number of anchors
              if (sheet.anchors && Array.isArray(sheet.anchors)) {
                console.log(`%c Anchors: ${sheet.anchors.length}`, 'color: #9b59b6');
                
                // Sample a few anchors if there are many
                const sampleSize = Math.min(5, sheet.anchors.length);
                if (sampleSize > 0) {
                  console.log('%c Sample anchors:', 'color: #9b59b6');
                  sheet.anchors.slice(0, sampleSize).forEach((anchor: any) => {
                    console.log(`  - ${anchor.address}: ${anchor.value || ''} ${anchor.type ? `(${anchor.type})` : ''}`);
                  });
                }
              }
              
              // Number of values
              if (sheet.values && Array.isArray(sheet.values)) {
                console.log(`%c Values: ${sheet.values.length}`, 'color: #3498db');
                
                // Sample a few values if there are many
                const sampleSize = Math.min(5, sheet.values.length);
                if (sampleSize > 0) {
                  console.log('%c Sample values:', 'color: #3498db');
                  sheet.values.slice(0, sampleSize).forEach((value: any) => {
                    console.log(`  - ${value.address}: ${value.value || ''} ${value.formula ? `(Formula: ${value.formula})` : ''}`);
                  });
                }
              }
              
              // Tables and charts
              if (sheet.tables && Array.isArray(sheet.tables)) {
                console.log(`%c Tables: ${sheet.tables.length}`, 'color: #f39c12');
              }
              
              if (sheet.charts && Array.isArray(sheet.charts)) {
                console.log(`%c Charts: ${sheet.charts.length}`, 'color: #e74c3c');
              }
              
              console.groupEnd(); // End sheet group
            });
            
            // Show raw JSON for developers who want to see everything
            console.groupCollapsed('%c Raw JSON Data', 'color: #7f8c8d');
            console.log(parsedContext);
            console.groupEnd();
            
            console.groupEnd(); // End main group
          } catch (error) {
            console.error('Error logging workbook context:', error);
            console.log('%c Raw workbook context:', 'color: #e74c3c');
            console.log(workbookContext.substring(0, 500) + '... [truncated]');
          }
        }
        
        // Prepare the messages array
        const messages = [];
        
        // Add system message
        messages.push({
          role: 'system',
          content: systemPrompt
        });
        
        // Add chat history for context
        if (chatHistory && chatHistory.length > 0) {
          for (const msg of chatHistory) {
            if (msg.attachments && msg.attachments.length > 0) {
              const content = [];
              
              // Add text content
              content.push({
                type: 'text',
                text: msg.content
              });
              
              // Add attachments
              for (const attachment of msg.attachments) {
                if (attachment.type === 'image') {
                  content.push({
                    type: 'image_url',
                    image_url: {
                      url: `data:${attachment.mimeType};base64,${attachment.content}`,
                      detail: 'high'
                    }
                  });
                } else if (attachment.type === 'pdf') {
                  content.push({
                    type: 'text',
                    text: `[Attached PDF: ${attachment.name}]`
                  });
                }
              }
              
              messages.push({
                role: msg.role as 'user' | 'assistant' | 'system',
                content: content
              });
            } else {
              messages.push({
                role: msg.role as 'user' | 'assistant' | 'system',
                content: msg.content
              });
            }
          }
        }
        
        // Add workbook context
        messages.push({
          role: 'user',
          content: `EXCEL WORKBOOK CONTEXT:\n${workbookContext}`
        });
        
        // Add user query with attachments if any
        if (attachments && attachments.length > 0) {
          const content = [];
          
          // Add text content
          content.push({
            type: 'text',
            text: userQuery
          });
          
          // Add image attachments
          for (const attachment of attachments) {
            if (attachment.type === 'image') {
              content.push({
                type: 'image_url',
                image_url: {
                  url: `data:${attachment.mimeType};base64,${attachment.content}`,
                  detail: 'high'
                }
              });
            } else if (attachment.type === 'pdf') {
              // For PDFs, add a note
              content.push({
                type: 'text',
                text: `[Attached PDF: ${attachment.name}]`
              });
            }
          }
          
          messages.push({
            role: 'user',
            content: content
          });
        } else {
          // Simple text-only message
          messages.push({
            role: 'user',
            content: userQuery
          });
        }
        
        // Handle streaming if requested
        if (streamHandler) {
          // Initialize variables to capture the streamed response
          fullResponse = '';
          
          // Create the streaming request
          const stream = await this.openai.chat.completions.create({
            model: modelToUse,
            messages: messages as any,
            max_tokens: 4000,
            temperature: 0.7,
            stream: true
          });
          
          // Process the stream
          for await (const chunk of stream) {
            const content = chunk.choices[0]?.delta?.content;
            if (content) {
              fullResponse += content;
              
              // Pass the chunk to the handler
              streamHandler(content);
            }
          }
          
          // Clear the timeout if the request completes successfully
          if (timeoutId) clearTimeout(timeoutId);
          
          return {
            id: uuidv4(),
            assistantMessage: fullResponse,
            command: null, // No commands for explanations
            rawResponse: undefined
          };
        } else {
          // For non-streaming responses
          response = await this.openai.chat.completions.create({
            model: modelToUse,
            messages: messages as any,
            max_tokens: 2000,
            temperature: 0.2,
          });
          
          // Extract message text from the response
          messageText = response.choices[0]?.message?.content || 'No response text received';
          
          // Clear the timeout if the request completes successfully
          if (timeoutId) clearTimeout(timeoutId);
          
          // Return the result
          return {
            id: uuidv4(),
            assistantMessage: messageText,
            command: null, // No commands for explanations
            rawResponse: this.debugMode ? response : undefined,
          };
        }
      } catch (error: any) {
        // Clear the timeout to prevent memory leaks
        if (timeoutId) clearTimeout(timeoutId);
        
        console.error(`Error generating workbook explanation (attempt ${retryCount + 1}/${MAX_RETRIES + 1}):`, error);
        lastError = error;
        
        // Check if the error is recoverable (rate limit, server error, etc.)
        if (this.isRecoverableError(error) && retryCount < MAX_RETRIES) {
          retryCount++;
          
          // Calculate exponential backoff delay: 2^retry * 1000ms + random jitter
          const backoffDelay = Math.min(
            (Math.pow(2, retryCount) * 1000) + (Math.random() * 1000),
            10000 // Cap at 10 seconds max
          );
          
          console.log(`Retrying in ${Math.round(backoffDelay / 1000)} seconds...`);
          await new Promise(resolve => setTimeout(resolve, backoffDelay));
          continue; // Try again
        }
        
        // If we're here, either we've exhausted retries or the error is not recoverable
        // Try to generate a simplified response with a smaller context if possible
        if (this.canUseReducedContext(error)) {
          try {
            console.log('Attempting to generate response with reduced context...');
            return await this.generateSimplifiedExplanation(userQuery, workbookContext, streamHandler);
          } catch (fallbackError) {
            console.error('Fallback explanation also failed:', fallbackError);
            // Continue to error handling below
          }
        }
        
        // Handle the error appropriately
        throw this.handleApiError(lastError);
      }
    }
    // This should never be reached, but TypeScript needs it for completeness
    throw new Error('Unexpected end of retry loop');
  }
  
  /**
   * Checks if an error is recoverable (can be retried)
   * @param error The error to check
   * @returns True if the error is recoverable, false otherwise
   */
  private isRecoverableError(error: any): boolean {
    // Check for rate limiting errors
    if (error.status === 429) return true;
    
    // Check for server errors (5xx)
    if (error.status >= 500 && error.status < 600) return true;
    
    // Check for specific OpenAI error types that are recoverable
    const errorType = error.error?.type;
    return [
      'rate_limit_exceeded',
      'server_error',
      'overloaded',
      'timeout'
    ].includes(errorType);
  }
  
  /**
   * Checks if we can use a reduced context approach for this error
   * @param error The error to check
   * @returns True if we can use reduced context, false otherwise
   */
  private canUseReducedContext(error: any): boolean {
    // Check for context length/token limit errors
    if (error.error?.type === 'context_length_exceeded') return true;
    if (error.error?.message?.includes('maximum context length')) return true;
    if (error.error?.message?.includes('token limit')) return true;
    if (error.error?.message?.includes('too many tokens')) return true;
    
    // Also try reduced context for timeout errors
    if (error.message?.includes('timed out')) return true;
    
    return false;
  }
  
  /**
   * Generate a simplified explanation with reduced context
   * @param userQuery The user's query
   * @param workbookContext The full workbook context
   * @param streamHandler Optional stream handler
   * @returns The simplified explanation
   */
  private async generateSimplifiedExplanation(
    userQuery: string,
    workbookContext: string,
    streamHandler?: (chunk: string) => void
  ): Promise<any> {
    try {
      // Parse the workbook context to reduce it
      const parsedContext = JSON.parse(workbookContext);
      
      // Create a simplified version with less detail
      const simplifiedContext = {
        activeSheet: parsedContext.activeSheet,
        sheets: parsedContext.sheets.map((sheet: any) => ({
          name: sheet.name,
          summary: sheet.summary,
          // Include only basic metadata about tables, charts, etc.
          tables: sheet.tables?.length ? `${sheet.tables.length} tables` : 'No tables',
          charts: sheet.charts?.length ? `${sheet.charts.length} charts` : 'No charts',
          // Limit the number of values and anchors
          anchors: sheet.anchors?.slice(0, 10).map((a: any) => ({ 
            address: a.address, 
            value: a.value, 
            type: a.type 
          })),
          values: sheet.values?.slice(0, 20).map((v: any) => ({ 
            address: v.address, 
            value: v.value,
            formula: v.formula
          }))
        }))
      };
      
      // Use a more concise system prompt
      const concisePrompt = `You are Cori, an Excel assistant. Analyze the simplified workbook data and answer the user's question concisely. Focus only on the most important aspects of the workbook. If you can't provide a detailed answer due to limited context, explain what you can determine and what information is missing.`;
      
      // Use a smaller model for faster response
      const modelToUse = this.models[ModelType.Light];
      
      // Create a simpler message structure
      const messages = [
        {
          role: 'system',
          content: concisePrompt
        },
        {
          role: 'user',
          content: `SIMPLIFIED EXCEL WORKBOOK CONTEXT:\n${JSON.stringify(simplifiedContext)}`
        },
        {
          role: 'user',
          content: userQuery
        }
      ];
      
      // Make the API call with reduced parameters
      if (streamHandler) {
        let fullResponse = '';
        const stream = await this.openai.chat.completions.create({
          model: modelToUse,
          messages: messages as any,
          max_tokens: 1000, // Reduced token limit
          temperature: 0.7,
          stream: true
        });
        
        for await (const chunk of stream) {
          const content = chunk.choices[0]?.delta?.content;
          if (content) {
            fullResponse += content;
            streamHandler(content);
          }
        }
        
        return {
          id: uuidv4(),
          assistantMessage: fullResponse,
          command: null,
          rawResponse: undefined
        };
      } else {
        const response = await this.openai.chat.completions.create({
          model: modelToUse,
          messages: messages as any,
          max_tokens: 1000, // Reduced token limit
          temperature: 0.2,
        });
        
        const messageText = response.choices[0]?.message?.content || 'No response text received';
        
        return {
          id: uuidv4(),
          assistantMessage: messageText,
          command: null,
          rawResponse: this.debugMode ? response : undefined,
        };
      }
    } catch (error) {
      console.error('Error generating simplified explanation:', error);
      throw error;
    }
  }
}