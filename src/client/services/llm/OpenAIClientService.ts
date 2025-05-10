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
// Import the ExcelCommandPlan and ExcelOperation types
import { ExcelCommandPlan, ExcelOperation } from '../../models/ExcelOperationModels';

export class OpenAIClientService {
  private openai: OpenAI;
  private debugMode: boolean = true;
  private verboseLogging: boolean = true;
  
  // Model configuration
  private models = {
    [ModelType.Light]: 'gpt-3.5-turbo',
    [ModelType.Standard]: 'gpt-4o-mini',
    [ModelType.Advanced]: 'gpt-4.1-nano-2025-04-14'
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
        const modelToUse = this.models[ModelType.Advanced];
        
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
        // comment this out for now
        console.log('chatHistory', chatHistory.slice(0, 0));
        /*
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
        */        
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
  
  /**
   * Generate Excel operations using OpenAI as a fallback when Anthropic fails
   * @param query User query for generating operations
   * @param workbookContext Context information about the workbook
   * @param chatHistory Previous chat history
   * @param attachments Optional image attachments
   * @param isRetry Whether this is a retry attempt after a failed parsing
   * @returns A plan with Excel operations
   */
  public async generateExcelOperations(
    query: string,
    workbookContext: string,
    chatHistory: Array<{ role: string; content: string }>,
    attachments?: Attachment[],
    isRetry: boolean = false
  ): Promise<ExcelCommandPlan> {
    try {
      // Parse the workbook context to extract formatting protocol if available
      let formattingProtocol = null;
      try {
        const parsedContext = JSON.parse(workbookContext);
        if (parsedContext.formattingProtocol) {
          formattingProtocol = parsedContext.formattingProtocol;
        }
      } catch (parseError) {
        console.warn('Error parsing workbook context to extract formatting protocol:', parseError);
      }
      
      // Build the system prompt
      const systemPrompt = this.buildExcelOperationsSystemPrompt(formattingProtocol, isRetry);
      
      // Filter the chat history to only include the last 5 messages
      const filteredChatHistory = chatHistory.slice(-5).filter(msg => msg.role !== 'system');
      const messageHistory = filteredChatHistory.filter(msg => msg.role === 'user' || msg.role === 'assistant');
      
      // Convert messageHistory to OpenAI message format
      const openaiMessages = messageHistory.map(msg => ({
        role: msg.role as 'user' | 'assistant',
        content: msg.content
      }));
      
      // Add the system prompt as the first message
      openaiMessages.unshift({
        role: 'system' as any, // Type assertion to handle OpenAI's message types
        content: systemPrompt
      });
      
      // Create the final user message
      const userPrompt = `User query: ${query}. Here is the workbook context to reference while generating operations: ${workbookContext}`;
      
      // Add attachments if they exist
      let finalUserContent: any = userPrompt; // Default to simple string content
      
      if (attachments && attachments.length > 0) {
        // For OpenAI, we need to format the content array differently than Anthropic
        const contentArray: any[] = [{ type: 'text', text: userPrompt }];
        
        for (const attachment of attachments) {
          if (attachment.type === 'image') {
            contentArray.push({
              type: 'image_url',
              image_url: {
                url: `data:${attachment.mimeType};base64,${attachment.content}`
              }
            });
          }
          // Note: OpenAI doesn't directly support PDF attachments like Anthropic does
        }
        
        // Add the multimodal message
        openaiMessages.push({
          role: 'user',
          content: contentArray as any // Type assertion for OpenAI's content types
        });
      } else {
        // Add text-only message
        openaiMessages.push({
          role: 'user',
          content: userPrompt
        });
      }
      
      // Use GPT-4 for better JSON generation
      const modelToUse = this.models[ModelType.Advanced];
      
      // Make the API call
      const response = await this.openai.chat.completions.create({
        model: modelToUse,
        messages: openaiMessages as any,
        max_tokens: 4000,
        temperature: isRetry ? 0.1 : 0.2, // Lower temperature for retry attempts
        response_format: { type: 'json_object' } // Enforce JSON response format
      });
      
      // Extract the response content
      const responseContent = response.choices[0]?.message?.content || '{"description":"Error generating operations","operations":[]}';
      
      try {
        // Parse the JSON response
        const plan = JSON.parse(responseContent) as ExcelCommandPlan;
        
        // Validate the operations
        this.validateOperations(plan.operations);
        
        return {
          description: plan.description || 'Excel operations',
          operations: plan.operations || []
        };
      } catch (parseError) {
        console.error('Failed to parse operations JSON from OpenAI:', parseError);
        
        // If this is not already a retry, try again with more explicit instructions
        if (!isRetry) {
          console.log('Retrying operation generation with OpenAI using explicit JSON instructions');
          return this.generateExcelOperations(query, workbookContext, chatHistory, attachments, true);
        }
        
        // If this is already a retry, return an empty plan
        return {
          description: 'Error parsing operations from OpenAI',
          operations: []
        };
      }
    } catch (error: any) {
      console.error('Error generating Excel operations with OpenAI:', error);
      return {
        description: 'Error generating operations with OpenAI',
        operations: []
      };
    }
  }
  
  /**
   * Build the system prompt for generating Excel operations
   * @param formattingProtocol Optional formatting protocol to include in the prompt
   * @param isRetry Whether this is a retry attempt after a failed parsing
   * @returns The system prompt
   */
  private buildExcelOperationsSystemPrompt(formattingProtocol: any = null, isRetry: boolean = false): string {
    let basePrompt = `You are an expert Excel assistant that generates operations for Excel workbooks. Your task is to analyze user queries and generate a list of Excel operations to fulfill their requests.

CRITICAL INSTRUCTION: ONLY generate operations that the user EXPLICITLY asks for. DO NOT add any additional operations that the user did not request. If the user asks to "add a new tab", ONLY create a new worksheet and DO NOT add any data, charts, or formatting to it unless specifically requested.`;
    
    // Add more explicit instructions for retry attempts
    if (isRetry) {
      basePrompt = `${basePrompt}

CRITICAL JSON FORMATTING INSTRUCTION: Your previous response contained invalid JSON that could not be parsed. You MUST respond with ONLY a valid JSON object. Do not include any explanations, notes, or text outside the JSON object. The JSON must exactly follow the schema provided below with no extra or missing fields. Ensure all quotes, brackets, and commas are properly placed.`;
    }
    
    // Add formatting protocol instructions if available
    if (formattingProtocol) {
      basePrompt += `

IMPORTANT: When generating operations that involve formatting or styling, follow the user's existing formatting conventions as described below. This ensures consistency with the user's workbook design.

FORMATTING PROTOCOL:
`;
      
      // Add color coding instructions
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
}`;
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
   * Classify a query and decompose it into steps
   * @param query The query to classify
   * @param chatHistory Optional chat history to provide context
   * @returns The classification result
   */

   public async classifyQueryAndDecompose(
      query: string,
      chatHistory: Array<{role: string, content: string, attachments?: Attachment[]}> = []
    ): Promise<any> {
      try {
        // Create a powerful system prompt for query classification and decomposition
        const systemPrompt = `You are a command classifier for an expert financial model assistant specialized in Excel workbooks. Your task is to analyze user queries, classify them, and decompose them into logical steps.
  The chat history is only for reference, you have to decompose only the MOST RECENT QUERY FROM THE USER. DON"T INCLUDE PRIOR QUERIES IN YOUR CLASSIFICATION. 

  CLASSIFICATION TYPES:
  - greeting: ONLY pure greetings or pleasantries with no task, question or command intent (like "hello", "how are you?", etc.)
  - workbook_question: Questions about the workbook that don't need KB access
  - workbook_command: Commands to modify the workbook that don't need KB access
  
  OUTPUT FORMAT:
  You must respond with a JSON object that follows this schema:
  {
    "query_type": string,  // The primary classification (one of the types above)
    "steps": [             // Array of steps to execute (can be one or multiple)
      {
        "step_index": number,       // 0-based index of this step in sequence
        "step_action": string,      // Short action description 
        "step_specific_query": string, // The specific sub-query for this step
        "step_type": string,        // Classification for this specific step
        "depends_on": number[]      // Indices of steps this step depends on (can be empty)
      }
    ]
  }
  
  EXAMPLES:
  
  Example 1 - Simple greeting:
  User: "Hi there, how are you today?"
  {
    "query_type": "greeting",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Respond to greeting",
        "step_specific_query": "Hi there, how are you today?",
        "step_type": "greeting",
        "depends_on": []
      }
    ]
  }
  
  Example 2 - Simple workbook question:
  User: "What's the total revenue in Q2 2023?"
  {
    "query_type": "workbook_question",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Analyze workbook for Q2 2023 revenue",
        "step_specific_query": "What's the total revenue in Q2 2023?",
        "step_type": "workbook_question",
        "depends_on": []
      }
    ]
  }
  
  Example 3 - KB-dependent question:
  User: "How does our Q1 performance compare to the industry benchmarks?"
  {
    "query_type": "workbook_question_kb",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Retrieve industry benchmarks from KB",
        "step_specific_query": "What are the industry benchmarks for Q1?",
        "step_type": "workbook_question_kb",
        "depends_on": []
      },
      {
        "step_index": 1,
        "step_action": "Compare workbook Q1 performance to benchmarks",
        "step_specific_query": "Compare our Q1 performance to the industry benchmarks",
        "step_type": "workbook_question",
        "depends_on": [0]
      }
    ]
  }
  
  Example 4 - Multi-step command:
  User: "Update the revenue projections for 2024 using a 5% growth rate and create a bar chart showing the quarterly breakdown"
  {
    "query_type": "workbook_command",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Update revenue projections with 5% growth",
        "step_specific_query": "Update the revenue projections for 2024 using a 5% growth rate",
        "step_type": "workbook_command",
        "depends_on": []
      },
      {
        "step_index": 1,
        "step_action": "Create quarterly revenue bar chart",
        "step_specific_query": "Create a bar chart showing the quarterly breakdown of 2024 revenue",
        "step_type": "workbook_command",
        "depends_on": [0]
      }
    ]
  }
  
  Example 5 - Command with KB dependency:
  User: "Create a new sheet with competitive analysis using data from our knowledge base"
  {
    "query_type": "workbook_command_kb",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Retrieve competitive data from KB",
        "step_specific_query": "Find competitive analysis data in knowledge base",
        "step_type": "workbook_question_kb",
        "depends_on": []
      },
      {
        "step_index": 1,
        "step_action": "Create new competitive analysis sheet",
        "step_specific_query": "Create a new sheet with the competitive analysis data",
        "step_type": "workbook_command",
        "depends_on": [0]
      }
    ]
  }

  Example 6 - Workbook Question:
  User: "What is the stock price calculated from the DCF and where do you find it?"
  
    {
    "query_type": "workbook_question",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Analyze workbook for stock price",
        "step_specific_query": "Determine the stock price calculated from the DCF analysis and its cell address",
        "step_type": "workbook_question",
        "depends_on": []
      }
    ]
  }
  Example 6 - Workbook Question:
  User: "what is the circular logic in the relationships between the income statement, balance sheet and cash flow statement"
  
    {
    "query_type": "workbook_question",
    "steps": [
      {
        "step_index": 0,
        "step_action": "Analyze workbook for circular logic",
        "step_specific_query": "Describe how the interdependencies among net income, cash balances, and equity create a circular relationship",
        "step_type": "workbook_question",
        "depends_on": []
      }
    ]
  }ample 7 - Workbook Question with KB:
  
  Important rules:
  1. Always use the most specific query_type that applies
  2. Decompose complex queries into logical steps
  3. Use the depends_on field to show dependencies between steps
  4. Keep step_action descriptions concise but clear
  5. Make step_specific_query suitable for direct execution by the appropriate handler
  6. IMPORTANT: Do NOT classify a query as a greeting if it contains a greeting word but also includes a question or command. For example, "Hi, what's the total revenue?" should be classified as workbook_question, not greeting
  7. Only classify as greeting if the SOLE intent is a greeting with no task
  8. Don't break down simple questions or commands into more than one step. Try to keep it to as few steps as possible. 

  ANTI-PATTERNS TO AVOID:
  1. DO NOT respond with "I'll analyze this query" or "I'll classify this as..."
  2. DO NOT include any explanatory text before or after the JSON
  4. DO NOT include phrases like "Here's the classification" or "Here's the JSON"
  5. DO NOT respond with "I understand you want me to..."
  6. DO NOT acknowledge the request in natural language
  7. NEVER respond with anything other than the raw JSON object
  8. NEVER create redundant steps that simply repeat prior steps. Don't be too atomic in your approach. Aim to condense workbook explanations into one step.

  
  YOUR ENTIRE RESPONSE MUST BE ONLY THE JSON OBJECT WITH NO OTHER TEXT.
  YOU ARE NOT RESPONDING TO A HUMAN. YOUR RESPONSE WILL ONLY BE SEEN BY AN INTERNAL PROCESSOR THAT EXPECTS RAW JSON.
  IF YOU ADD ANY TEXT BEFORE OR AFTER THE JSON, THE SYSTEM WILL BREAK.
  
  EXAMPLE OF CORRECT RESPONSE:
  {"query_type":"workbook_command","steps":[{"step_index":0,"step_action":"Update data","step_specific_query":"Update cell A1","step_type":"workbook_command","depends_on":[]}]}
  
  EXAMPLE OF INCORRECT RESPONSE:
  
  I'll classify this query for you. Here's the JSON:
  {"query_type":"workbook_command","steps":[{"step_index":0,"step_action":"Update data","step_specific_query":"Update cell A1","step_type":"workbook_command","depends_on":[]}]}`;
  
      // Use Sonnet for classification (most capable model)
      const modelToUse = this.models[ModelType.Advanced];
      
      // Prepare the messages array with chat history and the current query
      let messages = [];
      
      // First add a system message explaining the conversation context
      if (chatHistory.length > 0) {
        // Filter out system messages and limit to last 10 messages to avoid token limits
        // Anthropic API doesn't accept 'system' role in the messages array - only as a top-level parameter
        const recentHistory = chatHistory
          .filter(msg => msg.role !== 'system')
          .slice(-10);
        
        if (this.debugMode) {
          console.log('Filtered chat history:', recentHistory.length, 'of', chatHistory.length, 'messages');
        }
        
        messages = [...messages, ...recentHistory];
      }
      
      // Add the current query
      messages.push({
        role: 'user' as const,
        content: query
      });
      
      if (this.debugMode) {
        console.log('Chat history length:', chatHistory.length);
        console.log('Total messages for classification:', messages.length);
      }
      
      // For debugging
      if (this.debugMode) {
        console.log('Query Classification Request:', {
          model: modelToUse,
          query: query.substring(0, 50) + (query.length > 50 ? '...' : '')
        });
      }

      try {
        // Call the API to get the classification using OpenAI format
        const response = await this.openai.chat.completions.create({
          model: modelToUse,
          messages: [
            { role: 'system', content: systemPrompt },
            ...this.cleanMessagesForAPI(messages)
          ],
          max_tokens: 2000,
          temperature: 0.2, // Low temperature for more deterministic classification
          response_format: { type: 'json_object' } // Ensure JSON response
        });
        
        // Extract the classification result
        let responseContent = response.choices[0]?.message?.content || '{"query_type":"unknown","steps":[]}';
        
        if (this.debugMode) {
          console.log('Raw response content:', responseContent);
        }
        
        // Handle the response content properly
        let classification;
        
        try {
          // First try to parse it directly as JSON
          classification = JSON.parse(responseContent);
        } catch (parseError) {
          // If direct parsing fails, try to extract JSON from markdown
          try {
            // Extract JSON if it's wrapped in a markdown code block
            const extractedJson = this.extractJsonFromMarkdown(responseContent);
            
            // If we got a string back, parse it; otherwise use the extracted object
            classification = typeof extractedJson === 'string' 
              ? JSON.parse(extractedJson) 
              : extractedJson;
              
            if (!classification || typeof classification !== 'object') {
              // Fallback to a default classification if we couldn't parse anything
              classification = {"query_type":"workbook_question","steps":[{
                "step_index": 0,
                "step_action": "Answer workbook question",
                "step_specific_query": query,
                "step_type": "workbook_question",
                "depends_on": []
              }]};
            }
          } catch (extractError) {
            console.error('Error extracting JSON from response:', extractError);
            // Fallback to a default classification
            classification = {"query_type":"workbook_question","steps":[{
              "step_index": 0,
              "step_action": "Answer workbook question",
              "step_specific_query": query,
              "step_type": "workbook_question",
              "depends_on": []
            }]};
          }
        }
        
        if (this.debugMode) {
          console.log('Query Classification Result:', classification);
        }
        
        return classification;
        
      } catch (error) {
        console.error('Error classifying query:', error);
        throw this.handleApiError(error);
      }
    }
    catch (error) {
      console.error('Error classifying query:', error);
      throw this.handleApiError(error);
    }
  }

    /**
 * Clean and format messages for the OpenAI API
 * @param messages Array of message objects to clean
 * @returns Properly formatted messages for OpenAI API
 */
private cleanMessagesForAPI(messages: Array<{role: string, content: string, attachments?: Attachment[]}>): any[] {
  return messages.map(msg => {
    // Handle messages with attachments (for multimodal models)
    if (msg.attachments && msg.attachments.length > 0) {
      return {
        role: msg.role === 'assistant' ? 'assistant' : msg.role === 'system' ? 'system' : 'user',
        content: [
          { type: 'text', text: msg.content },
          ...msg.attachments.map(attachment => {
            if (attachment.type === 'image') {
              return {
                type: 'image_url',
                image_url: {
                  url: attachment.content,
                  detail: 'high'
                }
              };
            }
            return null;
          }).filter(Boolean)
        ]
      };
    }
    
    // Handle regular text messages
    return {
      role: msg.role === 'assistant' ? 'assistant' : msg.role === 'system' ? 'system' : 'user',
      content: msg.content
    };
  });
}
}