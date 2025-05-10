import { v4 as uuidv4 } from 'uuid';
import Anthropic from '@anthropic-ai/sdk';
import OpenAI from 'openai';
import { Mistral } from '@mistralai/mistralai';

interface Attachment {
  type: 'image' | 'pdf';
  name: string;
  mimeType: string;
  content: string;
}

// Model types for different query complexities
enum AnthropicModelType {
    Light = 'light', // For simple queries like greetings
    Standard = 'standard', // For regular workbook queries
    Advanced = 'advanced', // For complex operations
    Fast = 'fast' // For fast responses
}
// Model types for different query complexities
enum OpenAIModelType {
    Light = 'light', // For simple queries like greetings
    Standard = 'standard', // For regular workbook queries
    Advanced = 'advanced', // For complex operations
    Fast = 'fast' // For fast responses
}
enum MistralModelType {
    Light = 'light', // For simple queries like greetings
    Standard = 'standard', // For regular workbook queries
    Advanced = 'advanced', // For complex operations
    Fast = 'fast' // For fast responses
}

export class QAService {
  private anthropic: Anthropic;
  private openai: OpenAI;
  private mistral: Mistral;
  private anthropicModels: Record<AnthropicModelType, string>;
  private openaiModels: Record<OpenAIModelType, string>;
  private mistralModels: Record<MistralModelType, string>;
  private verboseLogging: boolean;
  private debugMode: boolean;
  private lastProvider: 'anthropic' | 'openai' | 'mistral' = 'openai'; // Start with OpenAI so first call will be Anthropic

  constructor(
    verboseLogging = false,
    debugMode = false,
  ) {
    // Initialize Anthropic client using API key from environment variables
    this.anthropic = new Anthropic({
      apiKey: process.env.ANTHROPIC_API_KEY || '',
      dangerouslyAllowBrowser: true // Enable browser usage
    });
    
    // Initialize OpenAI client using API key from environment variables
    this.openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY || '',
      dangerouslyAllowBrowser: true // For client-side use
    });
    
    this.mistral = new Mistral({
      apiKey: 'XHst9dbFgRSRoG5s4P9vawDk3Cn168tn',
    });
    
    this.anthropicModels = {
      [AnthropicModelType.Light]: 'claude-3-5-haiku-20241022',
      [AnthropicModelType.Standard]: 'claude-3-7-sonnet-20250219',
      [AnthropicModelType.Advanced]: 'claude-3-7-sonnet-20250219',
      [AnthropicModelType.Fast]: 'claude-3-5-haiku-20241022'
    }
    this.openaiModels = {
      [OpenAIModelType.Light]: 'gpt-4o-mini',
      [OpenAIModelType.Standard]: 'gpt-4.1-mini-2025-04-14',
      [OpenAIModelType.Advanced]: 'gpt-4.1-2025-04-14',
      [OpenAIModelType.Fast]: 'gpt-4.1-nano-2025-04-14'
    }
    this.mistralModels = {
      [MistralModelType.Light]: 'mistral-small-latest',
      [MistralModelType.Standard]: 'pixtral-12b-2409',
      [MistralModelType.Advanced]: 'pixtral-large-latest',
      [MistralModelType.Fast]: 'mistral-small-latest'
    }
    this.verboseLogging = verboseLogging;
    this.debugMode = debugMode;
  }

  /**
   * Determine if an error is recoverable (can be retried)
   * @param error The error to check
   * @returns True if the error is recoverable
   */
  private isRecoverableError(error: any): boolean {
    // Check for rate limit errors
    if (error.status === 429 || error.code === 'rate_limit_exceeded') {
      return true;
    }
    
    // Check for server errors (5xx)
    if (error.status >= 500 && error.status < 600) {
      return true;
    }
    
    // Check for timeout errors
    if (error.message && error.message.includes('timeout')) {
      return true;
    }
    
    // Check for network errors
    if (error.name === 'NetworkError' || 
        (error.message && error.message.includes('network'))) {
      return true;
    }
    
    return false;
  }

  /**
   * Handle API errors and format them for the client
   * @param error The error to handle
   * @returns A formatted error object
   */
  private handleApiError(error: any): Error {
    // If it's already a well-formatted error, just return it
    if (error instanceof Error) {
      return error;
    }
    
    // Handle Anthropic error format
    if (error.status && error.error) {
      return new Error(`Anthropic API error (${error.status}): ${error.error.message || JSON.stringify(error.error)}`);
    }
    
    // Handle OpenAI error format
    if (error.response && error.response.data) {
      return new Error(`OpenAI API error: ${error.response.data.error.message || JSON.stringify(error.response.data.error)}`);
    }
    
    // Generic error handling
    return new Error(`API error: ${error.message || JSON.stringify(error)}`);
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
    
    // Only retry once if needed
    const MAX_RETRIES = 3;
    let retryCount = 0;
    let lastError: any = null;
    
    // Always start with Anthropic, then rotate to OpenAI and Mistral
    // Initialize lastProvider to ensure we start with Anthropic
    this.lastProvider = 'anthropic';
    
    // Then set up the next provider in the rotation
    if (this.lastProvider === 'anthropic') {
      this.lastProvider = 'openai';
    } else if (this.lastProvider === 'openai') {
      this.lastProvider = 'mistral';
    } else {
      this.lastProvider = 'anthropic';
    }
    let currentProvider = this.lastProvider;
    
    console.log(`Using ${currentProvider} for this workbook explanation`);
    
    while (retryCount <= MAX_RETRIES) {
      try {
        // Variables to store response data
        let fullResponse = '';
        
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

        // Prepare the messages array for both providers
        const messages = [];
        
        if (this.verboseLogging) {
          console.log('chatHistory', chatHistory ? chatHistory.slice(0, 2) : 'No chat history');
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
            type: 'text' as const,
            text: userQuery
          });
          
          // Add image attachments
          for (const attachment of attachments) {
            if (attachment.type === 'image') {
              content.push({
                type: 'image' as const,
                source: {
                  type: 'base64' as const,
                  media_type: attachment.mimeType,
                  data: attachment.content
                }
              });
            } else if (attachment.type === 'pdf') {
              // For PDFs, add a note
              content.push({
                type: 'text' as const,
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

        // Try with the current provider
        if (currentProvider === 'anthropic') {
          return await this.callAnthropic(
            userQuery,
            systemPrompt,
            messages,
            streamHandler,
            timeoutId
          );
        } else if (currentProvider === 'mistral') {
          return await this.callMistral(
            userQuery,
            systemPrompt,
            messages,
            streamHandler,
            timeoutId
          );
        }
        else {
          return await this.callOpenAI(
            userQuery,
            systemPrompt,
            messages,
            streamHandler,
            timeoutId
          );
        }
      } catch (error: any) {
        // Clear the timeout to prevent memory leaks
        if (timeoutId) clearTimeout(timeoutId);
        
        console.error(`Error with ${currentProvider} (attempt ${retryCount + 1}/${MAX_RETRIES + 1}):`, error);
        lastError = error;
        
        // If we have a retry available, switch providers and try again
        if (retryCount < MAX_RETRIES) {
          retryCount++;
          
          // Switch providers for the retry
          currentProvider = currentProvider === 'anthropic' ? 'openai' : currentProvider === 'openai' ? 'mistral' : 'anthropic';
          console.log(`Switching to ${currentProvider} for retry...`);
          
          // Small delay before retry
          await new Promise(resolve => setTimeout(resolve, 500));
          continue; // Try again with the other provider
        }
        
        // If we're here, either we've exhausted retries or the error is not recoverable
        throw this.handleApiError(lastError);
      }
    }
    
    // This should never be reached, but TypeScript needs it for completeness
    throw new Error('Unexpected end of retry loop');
  }

  /**
   * Call Anthropic API for workbook explanation
   */
  private async callAnthropic(
    userQuery: string,
    systemPrompt: string,
    messages: any[],
    streamHandler?: (chunk: string) => void,
    timeoutId?: NodeJS.Timeout
  ): Promise<any> {
    // For workbook explanations, we'll use Sonnet which balances capability and speed
    const modelToUse = this.anthropicModels[AnthropicModelType.Advanced];
    
    // Ensure the user query is included
    if (!messages.some(msg => msg.content === userQuery || 
        (Array.isArray(msg.content) && msg.content.some(item => 
          item.type === 'text' && item.text === userQuery)))) {
      messages.push({
        role: 'user',
        content: userQuery
      });
    }
    
    // Handle streaming if requested
    if (streamHandler) {
      // Initialize variables to capture the streamed response
      let fullResponse = '';
      let isStreamingComplete = false;
      
      // Create a promise that resolves when streaming is complete
      const streamingCompletePromise = new Promise<string>((resolve) => {
        // Create the streaming request asynchronously
        (async () => {
          try {
            const stream = await this.anthropic.messages.create({
              model: modelToUse,
              system: systemPrompt,
              messages: messages as any,
              max_tokens: 4000,
              temperature: 0.7,
              stream: true
            });
            
            // Process the stream
            for await (const chunk of stream) {
              // Check for content block delta type
              if (chunk.type === 'content_block_delta') {
                // Safely access potentially text content
                const delta = chunk.delta as any; // Using any to bypass type checking for now
                if (delta && typeof delta.text === 'string') {
                  fullResponse += delta.text;
                  
                  // Pass the chunk to the handler
                  streamHandler(delta.text);
                }
              }
            }
            
            // Mark streaming as complete and resolve the promise
            isStreamingComplete = true;
            resolve(fullResponse);
          } catch (error) {
            console.error('Error during streaming:', error);
            resolve(fullResponse); // Resolve with what we have so far
          }
        })();
      });
      
      // Wait for streaming to complete
      const finalResponse = await streamingCompletePromise;
      
      // Clear the timeout if the request completes successfully
      if (timeoutId) clearTimeout(timeoutId);
      
      return {
        id: uuidv4(),
        assistantMessage: finalResponse,
        command: null, // No commands for explanations
        wasStreamed: true, // Flag to indicate content was already streamed
        rawResponse: undefined
      };
    } else {
      // For non-streaming responses
      const response = await this.anthropic.messages.create({
        model: modelToUse,
        system: systemPrompt,
        messages: messages as any,
        max_tokens: 4000,
        temperature: 0.7,
      });
      
      // Extract message text from the response
      const messageText = response.content?.[0]?.type === 'text' 
        ? response.content[0].text 
        : 'No response text received';
      
      // Clear the timeout if the request completes successfully
      if (timeoutId) clearTimeout(timeoutId);
      
      // Return the result
      return {
        id: uuidv4(),
        assistantMessage: messageText,
        command: null, // No commands for explanations
        wasStreamed: false, // Flag to indicate content was NOT streamed
        rawResponse: this.debugMode ? response : undefined,
      };
    }
  }

  /**
   * Call OpenAI API for workbook explanation
   */
  private async callOpenAI(
    userQuery: string,
    systemPrompt: string,
    messages: any[],
    streamHandler?: (chunk: string) => void,
    timeoutId?: NodeJS.Timeout
  ): Promise<any> {
    // Convert messages to OpenAI format
    const openaiMessages: Array<{
      role: 'system' | 'user' | 'assistant',
      content: string | Array<any>
    }> = [
      { role: 'system', content: systemPrompt }
    ];
    
    // Add the workbook context and user query
    for (const msg of messages) {
      if (typeof msg.content === 'string') {
        openaiMessages.push({
          role: msg.role as 'system' | 'user' | 'assistant',
          content: msg.content
        });
      } else if (Array.isArray(msg.content)) {
        // For multimodal content, we need to convert to OpenAI format
        const content: Array<any> = [];
        
        for (const item of msg.content) {
          if (item.type === 'text') {
            content.push({
              type: 'text',
              text: item.text
            });
          } else if (item.type === 'image') {
            content.push({
              type: 'image_url',
              image_url: {
                url: `data:${item.source.media_type};base64,${item.source.data}`,
                detail: 'high'
              }
            });
          }
        }
        
        openaiMessages.push({
          role: msg.role as 'system' | 'user' | 'assistant',
          content: content
        });
      }
    }
    
    // Ensure the user query is included
    if (!messages.some(msg => msg.content === userQuery || 
        (Array.isArray(msg.content) && msg.content.some(item => 
          item.type === 'text' && item.text === userQuery)))) {
      openaiMessages.push({
        role: 'user',
        content: userQuery
      });
    }
    
    // Handle streaming if requested
    if (streamHandler) {
      let fullResponse = '';
      let isStreamingComplete = false;
      
      // Create a promise that resolves when streaming is complete
      const streamingCompletePromise = new Promise<string>((resolve) => {
        // Create the streaming request asynchronously
        (async () => {
          try {
            const stream = await this.openai.chat.completions.create({
              model: this.openaiModels[OpenAIModelType.Fast],
              messages: openaiMessages,
              max_tokens: 4000,
              temperature: 0.2,
              stream: true
            });
            
            for await (const chunk of stream) {
              const content = chunk.choices[0]?.delta?.content || '';
              if (content) {
                fullResponse += content;
                streamHandler(content);
              }
            }
            
            // Mark streaming as complete and resolve the promise
            isStreamingComplete = true;
            resolve(fullResponse);
          } catch (error) {
            console.error('Error during streaming:', error);
            resolve(fullResponse); // Resolve with what we have so far
          }
        })();
      });
      
      // Wait for streaming to complete
      const finalResponse = await streamingCompletePromise;
      
      // Clear the timeout if the request completes successfully
      if (timeoutId) clearTimeout(timeoutId);
      
      return {
        id: uuidv4(),
        assistantMessage: finalResponse,
        command: null,
        wasStreamed: true, // Flag to indicate content was already streamed
        rawResponse: undefined
      };
    } else {
      const response = await this.openai.chat.completions.create({
        model: this.openaiModels[OpenAIModelType.Fast],
        messages: openaiMessages,
        max_tokens: 4000,
        temperature: 0.2
      });
      
      const messageText = response.choices[0]?.message?.content || 'No response received';
      
      // Clear the timeout if the request completes successfully
      if (timeoutId) clearTimeout(timeoutId);
      
      return {
        id: uuidv4(),
        assistantMessage: messageText,
        command: null,
        wasStreamed: false, // Flag to indicate content was NOT streamed
        rawResponse: this.debugMode ? response : undefined
      };
    }
  }

  /**
   * Call Mistral API for workbook explanation
   */
  private async callMistral(
  userQuery: string,
  systemPrompt: string,
  messages: any[],
  streamHandler?: (chunk: string) => void,
  timeoutId?: NodeJS.Timeout
): Promise<any> {
  // Convert messages to OpenAI format
  const openaiMessages: Array<{
    role: 'system' | 'user' | 'assistant',
    content: string | Array<any>
  }> = [
    { role: 'system', content: systemPrompt }
  ];
  
  // Add the workbook context and user query
  for (const msg of messages) {
    if (typeof msg.content === 'string') {
      openaiMessages.push({
        role: msg.role as 'system' | 'user' | 'assistant',
        content: msg.content
      });
    } else if (Array.isArray(msg.content)) {
      // For multimodal content, we need to convert to OpenAI format
      const content: Array<any> = [];
      
      for (const item of msg.content) {
        if (item.type === 'text') {
          content.push({
            type: 'text',
            text: item.text
          });
        } else if (item.type === 'image') {
          content.push({
            type: 'image_url',
            image_url: {
              url: `data:${item.source.media_type};base64,${item.source.data}`,
              detail: 'high'
            }
          });
        }
      }
      
      openaiMessages.push({
        role: msg.role as 'system' | 'user' | 'assistant',
        content: content
      });
    }
  }
  
  // Ensure the user query is included
  if (!messages.some(msg => msg.content === userQuery || 
      (Array.isArray(msg.content) && msg.content.some(item => 
        item.type === 'text' && item.text === userQuery)))) {
    openaiMessages.push({
      role: 'user',
      content: userQuery
    });
  }
  
  // Handle streaming if requested
  if (streamHandler) {
    let fullResponse = '';
    let isStreamingComplete = false;
    
    // Create a promise that resolves when streaming is complete
    const streamingCompletePromise = new Promise<string>((resolve) => {
      // Create the streaming request asynchronously
      (async () => {
        try {

            const result = await this.mistral.chat.stream({
                model: this.mistralModels[MistralModelType.Fast],
                messages: openaiMessages,
            });
          
          for await (const chunk of result) {
            const content = chunk.data.choices[0]?.delta?.content || '';
            const contentText = typeof content === 'string' ? content : content.join('');
            if (contentText) {
              fullResponse += contentText;
              streamHandler(contentText);
            }
          }
          
          // Mark streaming as complete and resolve the promise
          isStreamingComplete = true;
          resolve(fullResponse);
        } catch (error) {
          console.error('Error during streaming:', error);
          resolve(fullResponse); // Resolve with what we have so far
        }
      })();
    });
    
    // Wait for streaming to complete
    const finalResponse = await streamingCompletePromise;
    
    // Clear the timeout if the request completes successfully
    if (timeoutId) clearTimeout(timeoutId);
    
    return {
      id: uuidv4(),
      assistantMessage: finalResponse,
      command: null,
      wasStreamed: true, // Flag to indicate content was already streamed
      rawResponse: undefined
    };
  } else {
    const response = await this.mistral.chat.complete({
        model: "mistral-large-latest",
        messages: openaiMessages,
    });
    
    const messageText = response.choices[0]?.message?.content || 'No response received';
    
    // Clear the timeout if the request completes successfully
    if (timeoutId) clearTimeout(timeoutId);
    
    return {
      id: uuidv4(),
      assistantMessage: messageText,
      command: null,
      wasStreamed: false, // Flag to indicate content was NOT streamed
      rawResponse: this.debugMode ? response : undefined
    };
  }
}

}