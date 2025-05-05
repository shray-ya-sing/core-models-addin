import { v4 as uuidv4 } from 'uuid';
import Anthropic from '@anthropic-ai/sdk';
import { CommandStatus } from '../models/CommandModels';
import { ChatHistoryMessage } from './ClientQueryProcessor';

// First, define interfaces for the classification result
interface QueryStep {
  step_index: number;
  step_action: string;
  step_specific_query: string;
  step_type: string;
  depends_on: number[];
}

interface QueryClassification {
  query_type: string;
  steps: QueryStep[];
}


// Model types for different query complexities
export enum ModelType {
  Light = 'light', // For simple queries like greetings
  Standard = 'standard', // For regular workbook queries
  Advanced = 'advanced' // For complex operations
}

/**
 * Client-side service for interacting with the Anthropic API
 * Using the official Anthropic TypeScript SDK
 */
export class ClientAnthropicService {
  private anthropic: Anthropic;
  private debugMode: boolean = true;
  // Enable verbose logging of chunks sent to LLM (TEMPORARY)
  private verboseLogging: boolean = true;
  
  // Model configuration
  private models = {
    [ModelType.Light]: 'claude-3-5-haiku-20241022',     // Fast, efficient for simple tasks
    [ModelType.Standard]: 'claude-3-7-sonnet-20250219',  // Good balance
    [ModelType.Advanced]: 'claude-3-7-sonnet-20250219'   // Most powerful
  };

  // Default model selection
  private defaultModel: string = this.models[ModelType.Advanced];

  constructor(apiKey: string) {
    this.anthropic = new Anthropic({
      apiKey: apiKey,
      dangerouslyAllowBrowser: true // Enable browser usage
    });
  }
  
  /**
   * Get the model ID for a specific model type
   * @param modelType The type of model to use
   * @returns The model ID string
   */
  public getModel(modelType: ModelType = ModelType.Advanced): string {
    return this.models[modelType] || this.defaultModel;
  }
  
  /**
   * Get the Anthropic client instance
   * @returns The Anthropic client
   */
  public getClient(): Anthropic {
    return this.anthropic;
  }

  /**
   * Simple chat completion for basic queries like greetings
   * @param prompt The user's prompt
   * @param streamHandler Optional handler for streaming responses
   * @returns The generated response
   */
  public async generateChatResponse(
    prompt: string,
    streamHandler?: (chunk: string) => void
  ): Promise<any> {
    try {
      // Create a basic system prompt for simple chat interactions
      const systemPrompt = `Your name is Cori. You are a financial modeling assistant for Excel. You help users understand and modify their financial models.`;
      
      const messages = [{
        role: 'user' as const,
        content: `System: ${systemPrompt}\n\nUser: ${prompt}`
      }];
      
      // Always use the light model for simple chat completions
      const modelToUse = this.models[ModelType.Light];
      
      if (this.debugMode) {
        console.log('Simple Chat API Request:', {
          model: modelToUse,
          prompt: prompt.substring(0, 50) + (prompt.length > 50 ? '...' : '')
        });
      }
      
      // Check if we should stream the response
      if (streamHandler) {
        // For streaming responses
        let fullResponse = '';
        let responseId = uuidv4();
        
        const stream = await this.anthropic.messages.create({
          model: modelToUse,
          messages: messages as any, // Type assertion to resolve SDK type issue
          max_tokens: 1000,  // Shorter for chat responses
          temperature: 0.7,
          stream: true,
        });
        
        // Process the stream
        for await (const chunk of stream) {
          if (chunk.type === 'content_block_delta' && chunk.delta?.type === 'text_delta') {
            const textChunk = chunk.delta?.text || '';
            fullResponse += textChunk;
            streamHandler(textChunk);
          }
        }
        
        return {
          id: responseId,
          assistantMessage: fullResponse,
          command: null,  // No commands for simple chat
          rawResponse: null
        };
      } else {
        // For non-streaming responses
        const response = await this.anthropic.messages.create({
          model: modelToUse,
          messages: messages as any, // Type assertion to resolve SDK type issue
          max_tokens: 1000,
          temperature: 0.7,
        });
        
        // Extract message text from the response
        const messageText = response.content?.[0]?.type === 'text' 
          ? response.content[0].text 
          : 'No response text received';
        
        return {
          id: uuidv4(),
          assistantMessage: messageText,
          command: null,  // No commands for simple chat
          rawResponse: this.debugMode ? response : undefined,
        };
      }
    } catch (error: any) {
      console.error('Error generating chat response:', error);
      throw this.handleApiError(error);
    }
  }

  /**
   * Analyze a user query to classify its type and decompose it into logical steps
   * @param query The user query to analyze
   * @param chatHistory Previous conversation history for context
   * @returns Classification and decomposition of the query
   */
  public async classifyQueryAndDecompose(
    query: string,
    chatHistory: Array<{role: string, content: string}> = []
  ): Promise<any> {
    try {
      // Create a powerful system prompt for query classification and decomposition
      const systemPrompt = `Your name is Cori.You are an expert financial model assistant specialized in Excel workbooks. Your task is to analyze user queries, classify them, and decompose them into logical steps.

CLASSIFICATION TYPES:
- greeting: ONLY pure greetings or pleasantries with no task, question or command intent (like "hello", "how are you?", etc.)
- workbook_question: Questions about the workbook that don't need KB access
- workbook_question_kb: Questions about the workbook that require KB access
- workbook_command: Commands to modify the workbook that don't need KB access
- workbook_command_kb: Commands to modify the workbook that require KB access

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

Important rules:
1. Always use the most specific query_type that applies
2. Decompose complex queries into logical steps
3. Use the depends_on field to show dependencies between steps
4. Keep step_action descriptions concise but clear
5. Make step_specific_query suitable for direct execution by the appropriate handler
6. IMPORTANT: Do NOT classify a query as a greeting if it contains a greeting word but also includes a question or command. For example, "Hi, what's the total revenue?" should be classified as workbook_question, not greeting
7. Only classify as greeting if the SOLE intent is a greeting with no task`;

    // Use Sonnet for classification (most capable model)
    const modelToUse = this.models[ModelType.Standard];
    
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
    
    // Call the API to get the classification
    const response = await this.anthropic.messages.create({
      model: modelToUse,
      system: systemPrompt,
      messages: messages as any, // Type assertion to resolve SDK type issue
      max_tokens: 2000,
      temperature: 0.2, // Low temperature for more deterministic classification
    });
    
    // Extract the classification result
    let responseContent = response.content?.[0]?.type === 'text' 
      ? response.content[0].text 
      : '{"query_type":"unknown","steps":[]}';
    
    try {
      // Extract JSON if it's wrapped in a markdown code block
      responseContent = this.extractJsonFromMarkdown(responseContent);
      
      // Parse the JSON response
      const classification = JSON.parse(responseContent);
      
      if (this.debugMode) {
        console.log('Query Classification Result:', classification);
      }
      
      return classification;
    } catch (parseError) {
      console.error('Failed to parse classification JSON:', parseError);
      console.error('Raw response content:', responseContent);
      // Return a default classification if parsing fails
      return {
        query_type: 'unknown',
        steps: [{
          step_index: 0,
          step_action: 'Unknown action',
          step_specific_query: query,
          step_type: 'unknown',
          depends_on: []
        }]
      };
    }
  } catch (error: any) {
    console.error('Error classifying query:', error);
    throw this.handleApiError(error);
  }
}


  /**
   * Generate a stepwise plan response for complex Excel operations
   * @param prompt The user's prompt
   * @param context Information about the workbook (required for planning)
   * @param modelType Type of model to use (standard or advanced)
   * @param streamHandler Optional handler for streaming responses
   * @returns The generated response with command plan
   */
  public async generateResponse(
    prompt: string, 
    context?: any, 
    modelType: ModelType = ModelType.Advanced,
    streamHandler?: (chunk: string) => void
  ): Promise<any> {
    try {
      // Create the message payload
      const systemPrompt = `Your name is Cori.You are a financial modeling assistant for Excel. 
You help users understand and modify their financial models.
${context ? 'Here is information about the current workbook:' : ''}

Format your response using proper Markdown syntax:
- Use headings (## and ###) to organize your explanation
- Use bullet points or numbered lists where appropriate
- Use **bold** or *italic* for emphasis
- Use code formatting for formulas and Excel references: \`=SUM(A1:A10)\`
- Use tables for structured data where helpful

Ensure your response is well-structured with clear sections and formatting to maximize readability. BE AS CONCISE AS POSSIBLE. DO NOT REPEAT CONTENT OR ADD REDUNDANT INFORMATION.
RESPOND IN AS FEW CHARACTERS AS POSSIBLE.`;

      const messages = [];

      // Add context if provided
      if (context) {
        messages.push({
          role: 'user',
          content: `System: ${systemPrompt}\n\nWorkbook context: ${JSON.stringify(context)}\n\nUser: ${prompt}`
        });
      } else {
        messages.push({
          role: 'user',
          content: `System: ${systemPrompt}\n\nUser: ${prompt}`
        });
      }

      // Select the appropriate model based on query complexity
      const modelToUse = this.models[modelType] || this.defaultModel;
      
      // Log request details in debug mode
      if (this.debugMode) {
        console.log('Anthropic API Request:', {
          model: modelToUse,
          messagesCount: messages.length
        });
      }
      
      if (this.debugMode) {
        console.log(`Using model: ${modelToUse} for query type: ${modelType}`);
      }
      
      // Check if we should stream the response
      if (streamHandler) {
        // For streaming responses
        let fullResponse = '';
        let responseId = uuidv4();
        
        const stream = await this.anthropic.messages.create({
          model: modelToUse,
          messages: messages as any, // Type assertion to resolve SDK type issue
          max_tokens: 4000,
          temperature: 0.7,
          stream: true,
        });
        
        // Process the stream
        for await (const chunk of stream) {
          if (chunk.type === 'content_block_delta' && chunk.delta?.type === 'text_delta') {
            const textChunk = chunk.delta?.text || '';
            fullResponse += textChunk;
            streamHandler(textChunk);
          }
        }
        
        // Create a synthetic response object that matches the format of a non-streaming response
        return {
          id: responseId,
          assistantMessage: fullResponse,
          command: this.extractCommandPlan(fullResponse),
          rawResponse: null
        };
      } else {
        // For non-streaming responses
        const response = await this.anthropic.messages.create({
          model: modelToUse,
          messages: messages as any, // Type assertion to resolve SDK type issue
          max_tokens: 4000,
          temperature: 0.7,
        });
        
        if (this.debugMode) {
          console.log('Anthropic API Response received:', {
            contentLength: JSON.stringify(response).length,
            hasContent: !!response.content
          });
        }
        
        // Extract message text from the response
        const messageText = response.content?.[0]?.type === 'text' 
          ? response.content[0].text 
          : 'No response text received';
        
        // Extract any commands from the response
        const commandPlan = this.extractCommandPlan(messageText);
        
        // Return the result
        return {
          id: uuidv4(),
          assistantMessage: messageText,
          command: commandPlan,
          rawResponse: this.debugMode ? response : undefined,
        };
      }

    } catch (error: any) {
      console.error('Error generating stepwise plan:', error);
      throw this.handleApiError(error);
    }
  }
  

  /**
   * Extracts JSON content from a markdown code block if present
   * @param text The text that might contain a markdown code block
   * @returns The extracted JSON content or the original text if no code block is found
   */
  public extractJsonFromMarkdown(text: string): string {
    // Check if the text contains a code block
    const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)```/;
    const match = text.match(codeBlockRegex);
    
    if (match && match[1]) {
      const codeBlockContent = match[1].trim();
      
      if (this.debugMode) {
        console.log('Extracted JSON from markdown code block:', 
          codeBlockContent.substring(0, Math.min(50, codeBlockContent.length)) + 
          (codeBlockContent.length > 50 ? '...' : ''));
      }
      
      return codeBlockContent;
    }
    
    // Return the original text if it doesn't match code block pattern
    return text;
  }
  
  /**
   * Use LLM to select relevant sheets based on the user's query
   * @param query The user's query
   * @param availableSheets List of available sheets in the workbook
   * @returns Array of sheet names that are relevant to the query
   */
  public async selectRelevantSheets(
    query: string,
    availableSheets: Array<{name: string, summary: string}>,
    chatHistory: Array<{role: string, content: string}>
  ): Promise<string[]> {
    try {
      // Enhanced debug logging to track query through the method chain
      console.log(
        '%c ClientAnthropicService: Selecting sheets for query: ' + 
        `"${query}"`,
        'background: #27ae60; color: white; font-weight: bold; padding: 2px 5px;'
      );
      
      // Log query length and first/last characters to check for whitespace issues
      console.log(`%c Query length: ${query.length}, First char code: ${query.charCodeAt(0)}, Last char code: ${query.charCodeAt(query.length-1)}`, 'color: #2980b9;');
      
      // Log available sheets for debugging
      console.log(`%c Available sheets: ${availableSheets.map(s => s.name).join(', ')}`, 'color: #2980b9;');
      
      // Format the available sheets as a list
      const sheetsDescription = availableSheets.map(sheet => 
        `- "${sheet.name}": ${sheet.summary || 'No summary available'}`
      ).join('\n');
      
      // Create a clear, structured prompt for sheet selection
      const systemPrompt = `Your name is Cori. You are an expert Excel assistant that helps users find relevant sheets in their workbook.
      
YOUR TASK:
1. Given a user's query about an Excel workbook and a list of available sheets
2. Determine which sheets are most likely relevant to answering their query
3. Return ONLY a JSON array of sheet names, with no other text or explanation

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

IF IN DOUBT, include the sheet rather than exclude it. It's always better to include too many sheets than too few.

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

      // Create the message structure for the Anthropic API
      const messages = [
        {
          role: 'user' as const,
          content: [
            {
              type: 'text' as const,
              text: `${userPrompt}`
            }
          ]
        }
      ];
      
      // Use a fast model for this relatively simple task
      const modelToUse = this.models[ModelType.Light];
      
      // Make the API call
      const response = await this.anthropic.messages.create({
        model: modelToUse,
        messages: messages,
        max_tokens: 300,
        temperature: 0.1,  // Low temperature for consistent results
        system: systemPrompt
        // Note: We want JSON output, but we'll include this instruction in the prompt
        // as the response_format parameter isn't supported in this SDK version
      });
      
      // Extract and parse the JSON response
      const messageText = response.content?.[0]?.type === 'text' 
        ? response.content[0].text 
        : '[]';
      
      // Extract JSON array from the response
      const jsonText = this.extractJsonFromMarkdown(messageText);
      
      try {
        const result = JSON.parse(jsonText);
        
        // Check if the result contains a 'sheets' field or is a direct array
        const sheetNames = Array.isArray(result) ? result : 
                          (result.sheets && Array.isArray(result.sheets)) ? result.sheets : 
                          [];
        
        // Log the selected sheets
        if (this.debugMode) {
          console.log(`%c LLM selected sheets: ${sheetNames.join(', ')}`, 'color: #2ecc71');
        }
        
        return sheetNames;
      } catch (parseError) {
        console.error('Error parsing sheet selection response:', parseError);
        console.log('Raw response:', jsonText);
        return [];
      }
    } catch (error: any) {
      console.error('Error selecting relevant sheets:', error);
      // If there's an error, return empty array (fallback will happen elsewhere)
      return [];
    }
  }

  /**
   * Centralized error handling for API errors
   * @param error The error from the API
   * @returns A standardized error
   */
  private handleApiError(error: any): Error {
    if (error.status === 401) {
      return new Error(`Authentication error: The API key appears to be invalid or expired`); 
    } else if (error.status === 400) {
      return new Error(`Bad request: ${error.message || 'Unknown error'}`);
    } else if (error.status === 429) {
      return new Error(`Rate limit exceeded: Please try again in a few moments`);
    } else if (error.message?.includes('Failed to fetch')) {
      return new Error('Network error: Unable to connect to the Anthropic API. Please check your internet connection.');
    } else {
      return new Error(`Anthropic API error: ${error.message || 'Unknown error'}`);
    }
  }

  /**
   * Extract command plan from the assistant's response
   * @param responseText The assistant's response text
   * @returns The extracted command plan, or null if none found
   */
  private extractCommandPlan(responseText: string): any {
    try {
      // Look for special command markers in the response
      const commandRegex = /```json\s*([\s\S]*?)\s*```|\<command\>([\s\S]*?)\<\/command\>/i;
      const match = responseText.match(commandRegex);
      
      if (match) {
        const jsonContent = match[1] || match[2];
        try {
          const parsedCommand = JSON.parse(jsonContent);
          
          // Ensure steps array exists
          const steps = parsedCommand.steps || [];
          
          // Ensure each step has an operations array and required properties
          const validatedSteps = steps.map(step => ({
            ...step,
            description: step.description || 'Step',
            operations: Array.isArray(step.operations) ? step.operations : [],
            status: step.status || 'pending'
          }));
          
          // Return a properly formatted command
          return {
            id: uuidv4(),
            description: parsedCommand.description || 'Execute Excel operations',
            steps: validatedSteps,
            status: CommandStatus.Pending
          };
        } catch (parseError) {
          console.warn('Failed to parse command JSON:', parseError);
          return null;
        }
      }
      
      return null;
    } catch (error) {
      console.error('Error extracting command plan:', error);
      return null;
    }
  }
  
  /**
   * Generate a response to a workbook explanation query, using the workbook context
   * @param userQuery The user's query about the workbook
   * @param workbookContext The compressed workbook context
   * @param streamHandler Optional callback for handling streaming responses
   * @returns The generated response
   */
  async generateWorkbookExplanation(
    userQuery: string,
    workbookContext: string,
    streamHandler?: (chunk: string) => void,
    chatHistory?: Array<{role: string, content: string}>
  ): Promise<any> {
    try {
      // Create a system prompt specifically for workbook explanations
      const systemPrompt = `Your name is Cori.You are an Excel assistant that helps users understand and analyze their spreadsheets. 

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
      
      // For workbook explanations, we'll use Sonnet which balances capability and speed
      const modelToUse = this.models[ModelType.Standard];
      
      // Log the workbook context being sent to the LLM if verbose logging is enabled
      if (this.verboseLogging) {
        try {
          // Parse the JSON to format it nicely
          const parsedContext = JSON.parse(workbookContext);
          
          // Create a collapsible console group
          console.groupCollapsed(
            '%c 📊 WORKBOOK CHUNKS SENT TO LLM 📊',
            'background: #8e44ad; color: #ecf0f1; font-size: 14px; padding: 5px 10px; border-radius: 4px;'
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
      const messages = [
        {
          role: 'user',
          content: [
            {
              type: 'text',
              text: `EXCEL WORKBOOK CONTEXT:\n${workbookContext}\n\nQUESTION: ${userQuery}\n\nCHAT HISTORY:\n${chatHistory?.map(msg => `${msg.role}: ${msg.content}`).join('\n')}`
            }
          ]
        }
      ];
      
      // Handle streaming if requested
      if (streamHandler) {
        // Initialize variables to capture the streamed response
        let fullResponse = '';
        
        // Create the streaming request
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
        
        return {
          id: uuidv4(),
          assistantMessage: fullResponse,
          command: null, // No commands for explanations
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
        
        // Return the result
        return {
          id: uuidv4(),
          assistantMessage: messageText,
          command: null, // No commands for explanations
          rawResponse: this.debugMode ? response : undefined,
        };
      }
    } catch (error: any) {
      console.error('Error generating workbook explanation:', error);
      throw this.handleApiError(error);
    }
  }
}