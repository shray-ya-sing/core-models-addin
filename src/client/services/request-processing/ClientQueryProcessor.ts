// src/client/services/ClientQueryProcessor.ts
import { v4 as uuidv4 } from 'uuid';
import { ClientAnthropicService } from "../llm/ClientAnthropicService";
import { QAService } from "../qa/qaService";
import { ClientKnowledgeBaseService } from "../ClientKnowledgeBaseService";
import { ClientCommandManager } from "../actions/ClientCommandManager";
import { GreetingsService } from "../greetings/GreetingsService";
import { ClientExcelOperationGenerator } from '../actions/ClientExcelOperationGenerator';
import { ExcelCommandPlan } from '../../models/ExcelOperationModels';

import {
  Command,
  CommandStatus,
  QueryType as CommandModelQueryType,
  OperationType
} from '../../models/CommandModels';
import {
  ProcessStage,
  ProcessStatus,
  ProcessStatusManager
} from '../../models/ProcessStatusModels';
import { OpenAIClientService } from '../llm/OpenAIClientService';
import { LoadContextService } from '../context/LoadContextService';




/* ----------------------------  Query Typing  ---------------------------- */

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



export enum QueryType {
  Greeting = 'greeting',
  WorkbookQuestion = 'workbook_question',             // Needs workbook only
  WorkbookQuestionWithKB = 'workbook_question_kb',    // Workbook + KB
  WorkbookCommand = 'workbook_command',               // Command, no KB
  WorkbookCommandWithKB = 'workbook_command_kb',      // Command + KB
  Unknown = 'unknown'
}

export interface QueryProcessorResult {
  processId: string;
  assistantMessage: string;
  command?: Command | null;
}

// Interface for chat message history
export interface ChatHistoryMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
}

/* --------------------------  Main Class  -------------------------- */

export class ClientQueryProcessor {
  private anthropic: ClientAnthropicService;
  private kbService: ClientKnowledgeBaseService | null;
  private contextService: LoadContextService | null;
  private commandManager: ClientCommandManager | null;
  private currentQuerySteps: QueryStep[] = [];
  private openai?: OpenAIClientService;
  private useOpenAIFallback: boolean;
  // Greetings service for handling simple greetings
  private greetingsService: GreetingsService = new GreetingsService();
  private qaService: QAService;  

  constructor(params: {
    anthropic: ClientAnthropicService;
    kbService?: ClientKnowledgeBaseService | null;
    contextService?: LoadContextService | null;
    commandManager?: ClientCommandManager | null;
    openai?: OpenAIClientService;
    useOpenAIFallback?: boolean;
    qaService?: QAService;
  }) {
    this.anthropic = params.anthropic;
    this.kbService = params.kbService ?? null;
    this.contextService = params.contextService ?? null;
    this.commandManager = params.commandManager ?? null;
    this.currentQuerySteps = [];
    this.openai = params.openai;
    this.useOpenAIFallback = params.useOpenAIFallback !== undefined ? params.useOpenAIFallback : true;
    // Initialize QAService with verboseLogging and debugMode - it will load API keys from .env
    this.qaService = params.qaService ?? new QAService(false, false);
  }
  



  /* -----------------------  Public Entry Point  ----------------------- */

  async processQuery(
    userQuery: string,
    streamingCB?: (chunk: string) => void,
    chatHistory: Array<{role: string, content: string, attachments?: any[]}> = [],
    attachments?: any[]
  ): Promise<QueryProcessorResult> {
    const processId = uuidv4();
    const status = ProcessStatusManager.getInstance();
    let llmResponse = '';

    status.updateStatus(processId, {
      stage: ProcessStage.ResponseGeneration,
      status: ProcessStatus.Pending,
      message: 'Preparing internal resources'
    });
    
    console.group(`%c Query Processing (ID: ${processId.substring(0, 8)})`, 'background: #222; color: #f39c12; font-size: 14px; padding: 3px 6px;');
    console.log(`%c Received query: "${userQuery}"`, 'color: #2c3e50; font-size: 13px; font-weight: bold;');
    console.time('Query processing time');    

    status.updateStatus(processId, {
      stage: ProcessStage.ResponseGeneration,
      status: ProcessStatus.Pending,
      message: 'Generating response'
    });

    // Check if this is a simple greeting before doing expensive LLM operations
    if (this.greetingsService.isGreeting(userQuery) && streamingCB) {
      const response = this.greetingsService.handleGreeting(userQuery, streamingCB, processId);
      console.timeEnd('Query processing time');
      console.groupEnd();
      return {
        processId,
        assistantMessage: response
      };
    }
    
    // Set up progress update interval for UI responsiveness
    let progressUpdateCount = 0;
    const progressMessages = [
      "Let me think about that for a second before responding... I need to properly understand the workbook and its data",
      "\nI'm processing workbook data right now to understand how it's built out",
      "\nI'm examining worksheet structure for all the worksheets in the file...",
      "\nI'm retrieving relevant context...",
      "\nI'm preparing my final response...",
      "\nI'm sorry, this is taking me longer than expected... Give me a moment...",
      "\nI'm running into trouble analyzing the workbook. Please give me some more time to try again"
    ];
    
    // Set up tracking for active streaming
    let hasStartedStreaming = false;
    let lastStreamActivityTime = Date.now();
    
    // Create a wrapper for the streaming callback that tracks activity
    const streamingWrapper = streamingCB ? (chunk: string) => {
      // If this is an actual content chunk (not just a progress message)
      if (chunk.length > 5 && !chunk.startsWith('\n')) {
        hasStartedStreaming = true;
      }
      
      // Update the last activity timestamp
      lastStreamActivityTime = Date.now();
      
      // Call the original callback
      streamingCB(chunk);
    } : undefined;
  
    
    // Start progress updates if streaming is enabled
    let progressInterval: NodeJS.Timeout | null = null;
    if (streamingWrapper) {
      progressInterval = setInterval(() => {
        const message = progressMessages[progressUpdateCount % progressMessages.length];
        streamingWrapper(`\n${message}`);
        progressUpdateCount++;
      }, 2500); // Send a new message every 2.5 seconds
    }

    /* 1. Classify query with the LLM */
    const queryClassification = await this.classifyQueryAndDecompose(
      userQuery,
      processId,
      chatHistory
    );

    status.updateStatus(processId, {
      stage: ProcessStage.ResponseGeneration,
      status: ProcessStatus.Success,
      message: 'Generated plan'
    });

    if (!queryClassification) {
      console.log(`%c No steps found in the query classification result`, 'background: #c0392b; color: #fff; font-size: 12px; padding: 2px 5px;');
      console.timeEnd('Query processing time');
      console.groupEnd();
      return {
        processId,
        assistantMessage: "I couldn't process your request. Please try again with a clearer instruction."
      };
    }

    let combinedMessage = "";
    
    // Process steps sequentially with proper async/await
    for (let i = 0; i < queryClassification.steps.length; i++) {
      const step = queryClassification.steps[i];
      const queryType = step.step_type;
      const query = step.step_specific_query;
      const stepIndex = step.step_index;
      
      // If this isn't the first step and we have streaming enabled, send an inter-step message
      if (i > 0 && streamingWrapper) {
        streamingWrapper("\n\nI'm looking further... give me a moment to try something else...\n");
      }
      
      // Log that we're completing this step
      console.log(`%c Completing step ${stepIndex} of type ${queryType} with query: ${query}`, 'color: #2ecc71');

      try {
        // Handle different query types
        if (queryType === QueryType.Greeting) {
          // For greetings, no need to capture workbook state
          console.log(`%c Processing greeting: "${query}"`, `background: ${this.getClassificationColor(QueryType.Greeting)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
          
          // Clear the progress interval before sending the final response
          if (progressInterval) {
            clearInterval(progressInterval);
            progressInterval = null;
          }
          
          const response = this.greetingsService.handleGreeting(query, streamingWrapper, processId);
          console.timeEnd('Query processing time');
          console.groupEnd();
          return {
            processId,
            assistantMessage: response
          };
        } 
        else if (
          queryType === QueryType.WorkbookQuestion ||
          queryType === QueryType.WorkbookQuestionWithKB ||
          queryType === QueryType.WorkbookCommand ||
          queryType === QueryType.WorkbookCommandWithKB
        ) {
          // Use our contextBuilder to get chunks via the query-dependent selection
          const queryContext = await this.contextService.getQueryContextBuilder().buildContextForQuery(
            this.convertToCommandModelQueryType(queryType as QueryType),
            chatHistory,
            query
          );
          
          // Convert the context to JSON format
          const workbookJSON = this.contextService.getQueryContextBuilder().contextToJson(queryContext);
          
          // Route to appropriate handler based on query type
          if (queryType === QueryType.WorkbookQuestion) {
            console.log(`%c Processing workbook question: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookQuestion)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
            
            // Stop sending new progress messages
            if (progressInterval) {
              clearInterval(progressInterval);
              progressInterval = null;
            }
            
            const result = await this.answerWorkbookQuestion(
              query,
              workbookJSON,
              streamingWrapper,
              processId,
              chatHistory,
              attachments
            );
            combinedMessage += "\n"+result.assistantMessage;
          } else if (queryType === QueryType.WorkbookQuestionWithKB) {
            console.log(`%c Processing workbook question with KB: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookQuestionWithKB)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
            
            // Clear the progress interval before sending the final response
            if (progressInterval) {
              clearInterval(progressInterval);
              progressInterval = null;
            }
            
            const result = await this.answerWorkbookQuestionWithKB(
              query,
              workbookJSON,
              streamingWrapper,
              processId,
              chatHistory,
              attachments
            );  
            combinedMessage += "\n"+result.assistantMessage;
          } else if (queryType === QueryType.WorkbookCommand) {
            console.log(`%c Processing workbook command: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookCommand)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
            
            // Clear the progress interval before sending the final response
            if (progressInterval) {
              clearInterval(progressInterval);
              progressInterval = null;
            }
            
            const result = await this.runWorkbookCommand(
              query,
              workbookJSON,
              streamingWrapper,
              processId,
              chatHistory,
              attachments
            );
            combinedMessage+="\n"+result.assistantMessage;
          } else if (queryType === QueryType.WorkbookCommandWithKB) {
            console.log(`%c Processing workbook command with KB: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookCommandWithKB)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
            
            // Clear the progress interval before sending the final response
            if (progressInterval) {
              clearInterval(progressInterval);
              progressInterval = null;
            }
            
            const result = await this.runWorkbookCommand(
              query,
              workbookJSON,
              streamingWrapper,
              processId,
              chatHistory,
              attachments
            );
            combinedMessage+="\n"+result.assistantMessage;
          }
        }
        else if (queryType === QueryType.Unknown) {
          console.log(`%c Unable to classify query: "${query}"`, `background: ${this.getClassificationColor(QueryType.Unknown)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
          
          // Clear the progress interval before sending the final response
          if (progressInterval) {
            clearInterval(progressInterval);
            progressInterval = null;
          }
          
          const result = {
            processId,
            assistantMessage: "I'm sorry, I'm not sure how to handle that request yet."
          };
          
        }
      } catch (error) {
        console.error(`Error processing step ${stepIndex}:`, error);
        
        // Clear the progress interval and execution timeout
        if (progressInterval) {
          clearInterval(progressInterval);
          progressInterval = null;
        }

        return {
          processId,
          assistantMessage: combinedMessage,
          command: null
        };
      }
    }
    
    // Clear the progress interval and execution timeout when processing is complete
    if (progressInterval) {
      clearInterval(progressInterval);
      progressInterval = null;
    }
    
    // After processing all steps, return a combined result
    console.log(`%c All steps processed successfully`, 'background: #27ae60; color: #fff; font-size: 12px; padding: 2px 5px;');
    console.timeEnd('Query processing time');
    console.groupEnd();

    // If we get here, we should have a valid result from one of the handlers
    // This is a fallback in case none of the handlers returned a result
    
    // Make sure progress interval is cleared
    if (progressInterval) {
      clearInterval(progressInterval);
      progressInterval = null;
    }
    
    return {
      processId,
      assistantMessage: combinedMessage,
      command: null
    };
  }

  /* --------------------------- Helper Methods --------------------------- */

  /**
   * Generate a dynamic progress message using the LLM based on the user's query
   * This provides a more natural and engaging experience while waiting for the main response
   * @param query The user's original query
   * @param stage The current processing stage
   * @param streamHandler Optional handler for streaming the progress message
   * @returns Promise with the generated progress message
   */
  private async generateDynamicProgressMessage(query: string, stage: string, streamHandler?: (chunk: string) => void): Promise<string> {
    try {
      // If we have a stream handler, use streaming
      if (streamHandler) {
        // Start with a newline to separate from previous content
        streamHandler('\n');
        
        // Use a lightweight model for quick generation with streaming
        const response = await this.anthropic.generateQuickResponse(
          `Generate a natural, conversational progress message as if you're thinking out loud while working on answering this query: "${query}". 
          Current stage: ${stage}. 
          Keep it brief (under 100 characters), friendly, and make it sound like you're actively working. 
          Don't actually answer the query, just acknowledge you're working on it. 
          Don't use quotation marks in your response.`,
          { temperature: 0.7, max_tokens: 100 },
          // Pass a custom stream handler that prepends a space to each chunk
          (chunk) => streamHandler(chunk)
        );
        
        return '\n' + response.assistantMessage;
      } else {
        // Non-streaming fallback
        const response = await this.anthropic.generateQuickResponse(
          `Generate a natural, conversational progress message as if you're thinking out loud while working on answering this query: "${query}". 
          Current stage: ${stage}. 
          Keep it brief (under 100 characters), friendly, and make it sound like you're actively working. 
          Don't actually answer the query, just acknowledge you're working on it. 
          Don't use quotation marks in your response.`,
          { temperature: 0.7, max_tokens: 100 }
        );
        
        return '\n' + response.assistantMessage;
      }
    } catch (error) {
      console.error('Error generating dynamic progress message:', error);
      // Fallback to static messages if LLM generation fails
      const fallbackMessages = [
        "\nI'm analyzing your Excel workbook...",
        "\nLooking at the data structure and relationships...",
        "\nExamining the formulas and connections between sheets...",
        "\nProcessing your request, this will just take a moment..."
      ];
      const message = fallbackMessages[Math.floor(Math.random() * fallbackMessages.length)];
      
      // If we have a stream handler, send the fallback message through it
      if (streamHandler) {
        streamHandler(message);
      }
      
      return message;
    }
  }
 
  private getCurrentQueryType(pid: string): QueryType {
    // Get the query type from process status
    const sm = ProcessStatusManager.getInstance();
    const process = sm.getProcess(pid);
    
    if (!process || !process.data || !process.data.queryType) {
      return QueryType.Unknown;
    }
    
    return process.data.queryType as QueryType;
  }
  
  private getCurrentQueryText(pid: string): string {
    // Get the query text from process status
    const sm = ProcessStatusManager.getInstance();
    const process = sm.getProcess(pid);
    
    if (!process || !process.data || !process.data.query) {
      return '';
    }
    
    return process.data.query as string;
  }
  
  /**
   * Convert from local QueryType to CommandModels QueryType
   * @param localType The local QueryType enum value
   * @returns The equivalent CommandModels QueryType enum value
   */
  private convertToCommandModelQueryType(localType: QueryType): CommandModelQueryType {
    // Map from local QueryType to CommandModels QueryType
    switch (localType) {
      case QueryType.Greeting:
        return CommandModelQueryType.Greeting;
      case QueryType.WorkbookQuestion:
        return CommandModelQueryType.WorkbookQuestion;
      case QueryType.WorkbookQuestionWithKB:
        return CommandModelQueryType.WorkbookQuestionWithKB;
      case QueryType.WorkbookCommand:
        return CommandModelQueryType.WorkbookCommand;
      case QueryType.WorkbookCommandWithKB:
        return CommandModelQueryType.WorkbookCommandWithKB;
      default:
        return CommandModelQueryType.Unknown;
    }
  }

  private async classifyQuery(
    query: string,
    chatHistory: Array<{role: string, content: string, attachments?: any[]}> = []
  ): Promise<QueryClassification> {
    try {
      const res = await this.anthropic.classifyQueryAndDecompose(query, chatHistory);

      // Deserialize the classification result
      const classification = res as QueryClassification;
      
      return classification;
    } catch (anthropicError) {
      console.warn('Anthropic classification failed, falling back to OpenAI:', anthropicError);
      
      // If Anthropic fails, try OpenAI as fallback
      try {
        const res = await this.openai.classifyQueryAndDecompose(query, chatHistory);
        
        // Deserialize the classification result
        const classification = res as QueryClassification;
        
        return classification;
      } catch (openaiError) {
        console.error('OpenAI classification failed, falling back to simple explanation:', openaiError);
        return {
          query_type: 'workbook_question',
        steps: [
          {
            step_index: 0,
            step_action: 'Answer question about workbook',
            step_specific_query: query,
            step_type: 'workbook_question',
            depends_on: []
          }
        ]
      };
    }
  }
}

  private async classifyQueryAndDecompose(
    query: string,
    _pid: string, 
    chatHistory: Array<{role: string, content: string, attachments?: any[]}> = []
  ): Promise<QueryClassification> {
    try {
      const res = await this.classifyQuery(query, chatHistory);
      
      // Deserialize the classification result
      const classification = res as QueryClassification;
      
      // Store the steps for later use in a class property or state manager
      if (classification.steps && classification.steps.length > 0) {
        // Store all steps
        this.currentQuerySteps = classification.steps;
        
        // Enhanced logging with styled console logs for better visibility
        console.group('%c Query Decomposition Results', 'background: #222; color: #fff; font-size: 14px; padding: 2px 5px;');
        console.log(`Original query: "${query}"`);
        console.log(`%c Primary classification: ${classification.query_type}`, 
          `background: ${this.getClassificationColor(classification.query_type)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
        console.log(`%c Decomposed into ${classification.steps.length} step(s)`, 'background: #222; color: #3498db; font-size: 12px; padding: 2px 5px;');
        
        classification.steps.forEach(step => {
          console.group(`%c Step ${step.step_index}: ${step.step_action}`, 
            'background: #34495e; color: #fff; font-size: 13px; padding: 2px 5px;');
          console.log(`%c Type: ${step.step_type}`, 
            `background: ${this.getClassificationColor(step.step_type)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
          console.log(`Sub-query: "${step.step_specific_query}"`);
          if (step.depends_on && step.depends_on.length > 0) {
            console.log(`%c Dependencies: ${step.depends_on.join(', ')}`, 'color: #e74c3c; font-weight: bold;');
          } else {
            console.log('%c Dependencies: None', 'color: #7f8c8d;');
          }
          console.groupEnd();
        });
        
        console.groupEnd();
      }
      
      // Return the primary query type for routing
      return classification;
    } catch (err) {
      console.error(err);
      // Return a default classification on error
      return {
        query_type: QueryType.WorkbookQuestion,
        steps: [{
          step_index: 0,
          step_action: "Process query",
          step_specific_query: query,
          step_type: QueryType.WorkbookQuestion,
          depends_on: []
        }]
      };
    }
  }


  private getClassificationColor(queryType: string): string {
    switch (queryType) {
      case QueryType.Greeting:
        return '#2ecc71'; // Green
      case QueryType.WorkbookQuestion:
        return '#3498db'; // Blue
      case QueryType.WorkbookQuestionWithKB:
        return '#9b59b6'; // Purple
      case QueryType.WorkbookCommand:
        return '#e67e22'; // Orange
      case QueryType.WorkbookCommandWithKB:
        return '#e74c3c'; // Red
      case QueryType.Unknown:
      default:
        return '#95a5a6'; // Gray
    }
  }

    /* -----------------------  Individual Handlers  ----------------------- */

    private async handleGreeting(
      query: string,
      stream: ((c: string) => void) | undefined,
      pid: string,
      attachments?: any[]
    ): Promise<QueryProcessorResult> {
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.QueryProcessing,
        status: ProcessStatus.Pending,
        message: 'Generating greeting…'
      });
  
      const response = await this.anthropic.generateChatResponse(query, attachments, stream);
  
      sm.updateStatus(pid, {
        stage: ProcessStage.QueryProcessing,
        status: ProcessStatus.Success,
        message: 'Greeting sent.'
      });
  
      return { processId: pid, assistantMessage: response.assistantMessage };
    }
  
  
    private async answerWorkbookQuestion(
      query: string,
      workbookJSON: string,
      stream: ((c: string) => void) | undefined,
      pid: string,
      chatHistory: Array<{role: string, content: string, attachments?: any[]}>,
      attachments?: any[]
    ): Promise<QueryProcessorResult> {
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.ResponseGeneration,
        status: ProcessStatus.Pending,
        message: 'Analyzing sheets'
      });
      
      // Flag to track if the actual LLM response has started
      let actualResponseStarted = false;
      
      // Create a modified stream handler that stops placeholder text when actual response starts
      const enhancedStreamHandler = stream ? 
        (chunk: string) => {
          // If this is the first chunk from the actual LLM response
          if (!actualResponseStarted) {
            actualResponseStarted = true;
            // Add a line break to separate from any placeholder text
            stream('\n\n');
          }
          // Pass the chunk to the original stream handler
          stream(chunk);
        } : undefined;
      
      // Start sending immediate feedback messages if streaming is enabled
      let feedbackInterval: NodeJS.Timeout | null = null;
      if (stream) {
        // Send the first immediate feedback message
        stream('\nI\'m analyzing your Excel workbook... ');
        
        // Set up a sequence of feedback messages to show while waiting
        const feedbackMessages = [
          '\nI\'m looking at the workbook structure and examining any linkages across tabs...',
          '\nI\'m examining the data patterns and relationships to determine how the data is structured... ',
          '\nI\'m analyzing formulas and relationships to understand how the data values are being calculated... ',
          '\nI\'m identifying key insights and relationships to fully understand the logic behind the analysis... '
        ];
        
        let messageIndex = 0;
        // Send a new message every 2-3 seconds to show progress
        feedbackInterval = setInterval(() => {
          // Stop sending placeholder messages if actual response has started
          if (actualResponseStarted) {
            clearInterval(feedbackInterval!);
            feedbackInterval = null;
            return;
          }
          
          if (messageIndex < feedbackMessages.length) {
            stream(feedbackMessages[messageIndex]);
            messageIndex++;
          }
        }, 2500);
      }
      
      // Recovery messages to show if the API call fails
      const recoveryMessages = [
        "I'm having some trouble connecting to the server...",
        "There seems to be a network issue. I'll keep trying...",
        "The server is currently busy. Please bear with me...",
        "I'm experiencing some connectivity problems right now..."
      ];
      
      try {
        // Call the actual workbook explanation generation
        const response = await this.qaService.generateWorkbookExplanation(
          query,
          workbookJSON,
          enhancedStreamHandler,
          chatHistory,
          attachments
        );
        
        // Clear the feedback interval if it exists
        if (feedbackInterval) {
          clearInterval(feedbackInterval);
        }
        
        sm.updateStatus(pid, {
          stage: ProcessStage.ResponseGeneration,
          status: ProcessStatus.Success,
          message: 'Successfully analyzed workbook'
        });
        
        // For streamed responses, we still want to return the complete message
        // The streaming has already completed at this point
        return {
          processId: pid,
          assistantMessage: response.assistantMessage,
          command: null
        };
      } catch (error) {
        console.error('Error in workbook explanation with Anthropic:', error);
        
        // Clear the feedback interval if it exists
        if (feedbackInterval) {
          clearInterval(feedbackInterval);
        }
        
        // If there was an error and we're streaming, provide a message about trying an alternative service
        if (stream) {
          // Add a line break to separate from previous messages
          stream('\n\n');
          stream('I\'m having some trouble connecting to our primary service. Let me try an alternative approach...');
        }
        
        // Update status to show we're trying an alternative service
        sm.updateStatus(pid, {
          stage: ProcessStage.ResponseGeneration,
          status: ProcessStatus.Error,
          message: 'Could not complete request'
        })

        return {
          processId: pid,
          assistantMessage: "I'm sorry, I couldn't complete your request at this time. Please try again later.",
          command: null
        };
      }
    }
  
    /**
     * Creates an enhanced stream handler that manages the transition from feedback messages to actual response
     * @param originalStreamHandler The original stream handler function
     * @param query The user's query
     * @returns A new stream handler function
     */
    private createEnhancedStreamHandler(originalStreamHandler: (chunk: string) => void): (chunk: string) => void {
      let isFirstResponseChunk = true;
      
      return (chunk: string) => {
        // For the first chunk from the actual LLM response, add a line break and a clear indicator
        if (isFirstResponseChunk) {
          originalStreamHandler('\n\nHere\'s my analysis:\n');
          isFirstResponseChunk = false;
        }
        
        // Pass the chunk to the original handler
        originalStreamHandler(chunk);
      };
    }
    
    private async answerWorkbookQuestionWithKB(
      query: string,
      workbookJSON: string,
      stream: ((c: string) => void) | undefined,
      pid: string,
      chatHistory: Array<{role: string, content: string, attachments?: any[]}>,
      attachments?: any[]
    ): Promise<QueryProcessorResult> {
      /* 1. Fetch KB */
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.KnowledgeBaseQuery,
        status: ProcessStatus.Pending,
        message: 'Retrieving KB…'
      });
  
      try {
        const kbResults = this.kbService
          ? await this.kbService.search(query)
          : [];
  
        sm.updateStatus(pid, {
          stage: ProcessStage.KnowledgeBaseQuery,
          status: ProcessStatus.Success,
          message: `Found ${kbResults.length} KB items.`
        });
  
        /* 2. Append KB to workbook context and answer */
        const context = JSON.stringify({
          workbook: JSON.parse(workbookJSON),
          kb: kbResults
        });
  
        return this.answerWorkbookQuestion(query, context, stream, pid, chatHistory, attachments);
      } catch (error) {
        sm.updateStatus(pid, {
          stage: ProcessStage.KnowledgeBaseQuery,
          status: ProcessStatus.Error,
          message: `Error fetching KB: ${error}`
        });
  
        return {
          processId: pid,
          assistantMessage:
            "I'm sorry, I encountered an error while fetching knowledge base information. Please try again."
        };
      }
    }
  
    private async runWorkbookCommand(
      query: string,
      workbookJSON: string,
      stream: ((c: string) => void) | undefined,
      pid: string,
      chatHistory: Array<{ role: string; content: string; attachments?: any[] }>,
      attachments?: any[]
    ): Promise<QueryProcessorResult> {
      /* 1. Initialize status tracking */
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.CommandPlanning,
        status: ProcessStatus.Pending,
        message: 'Planning actions'
      });

      // Flag to track if the actual LLM response has started
      let actualResponseStarted = false;
      
      // Create a modified stream handler that stops placeholder text when actual response starts
      const enhancedStreamHandler = stream ? 
        (chunk: string) => {
          // If this is the first chunk from the actual LLM response
          if (!actualResponseStarted) {
            actualResponseStarted = true;
            // Add a line break to separate from any placeholder text
            stream('\n\n');
          }
          // Pass the chunk to the original stream handler
          stream(chunk);
        } : undefined;
      
      // Start sending immediate feedback messages if streaming is enabled
      let feedbackInterval: NodeJS.Timeout | null = null;
      if (stream) {
        // Send the first immediate feedback message
        stream('\nI\'m analyzing your Excel workbook... ');
        
        // Set up a sequence of feedback messages to show while waiting
        const feedbackMessages = [
          '\nI\'m planning how to execute this request for you...',
          '\nI\'m gathering internal resources to execute this request... ',
          '\nI\'m retriveing the right tools to help me execute this request... ',
          '\nI\'m Identifying the right cells I need to execute this request... '
        ];
        
        let messageIndex = 0;
        // Send a new message every 2-3 seconds to show progress
        feedbackInterval = setInterval(() => {
          // Stop sending placeholder messages if actual response has started
          if (actualResponseStarted) {
            clearInterval(feedbackInterval!);
            feedbackInterval = null;
            return;
          }
          
          if (messageIndex < feedbackMessages.length) {
            stream(feedbackMessages[messageIndex]);
            messageIndex++;
          }
        }, 2500);
      }

      // Recovery messages to show if the API call fails
      const recoveryMessages = [
        "I'm having some trouble connecting to the server...",
        "There seems to be a network issue. I'll keep trying...",
        "The server is currently busy. Please bear with me...",
        "I'm experiencing some connectivity problems right now..."
      ];
    
      try {
        // Ensure commandManager is available
        if (!this.commandManager) {
          throw new Error('Command manager is not initialized');
        }
    
        // Create the Excel operation generator
        const operationGenerator = new ClientExcelOperationGenerator({
          anthropic: this.anthropic,
          debugMode: true
        });
    
        // Generate operations based on the query and workbook context
        const operationPlan = await operationGenerator.generateOperationsWithMultimodal(
          query,
          workbookJSON,
          chatHistory,
          attachments
        );
    
        // Create a Command object from the operation plan
        const command: Command = {
          id: uuidv4(),
          description: operationPlan.description,
          status: CommandStatus.Pending,
          createdAt: new Date(),
          steps: [{
            description: operationPlan.description,
            operations: operationPlan.operations.map(op => {
              // Determine the appropriate target based on operation type
              let target = 'workbook'; // Default fallback
              
              // Extract target from operation where available
              if ('target' in op) {
                target = op.target;
              } else if ('range' in op) {
                target = op.range;
              }               
              return {
                type: op.op as unknown as OperationType,
                target: target,
                value: op
              };
            }),
            status: CommandStatus.Pending
          }]
        };
        
        // Update status to show we're ready to execute
        sm.updateStatus(pid, {
          stage: ProcessStage.CommandPlanning,
          status: ProcessStatus.Success,
          message: 'Preparing to execute...'
        });
    
        // Add command to the manager and register for status updates
        this.commandManager.addCommand(command);
        
        // Set up command status monitoring
        let commandCompleted = false;
        const statusListener = (updatedCommand: Command) => {
          if (updatedCommand.id !== command.id) return;
          
          if (updatedCommand.status === CommandStatus.Running) {
            sm.updateStatus(pid, {
              stage: ProcessStage.CommandExecution,
              status: ProcessStatus.Pending,
              message: `Processing your request...`,
            });
          } else if (updatedCommand.status === CommandStatus.Completed) {
            sm.updateStatus(pid, {
              stage: ProcessStage.CommandExecution,
              status: ProcessStatus.Success,
              message: 'I completed the operation successfully',
            });
            commandCompleted = true;
            unregisterListener(); // Call the function returned by onCommandUpdate
          } else if (updatedCommand.status === CommandStatus.Failed) {
            sm.updateStatus(pid, {
              stage: ProcessStage.CommandExecution,
              status: ProcessStatus.Error,
              message: `I failed to complete the operation. Please try again`,
            });
            commandCompleted = true;
            unregisterListener(); // Call the function returned by onCommandUpdate
          }
        };

        // Register for updates - this returns an unregister function
        const unregisterListener = this.commandManager.onCommandUpdate(statusListener);

        // Execute the command asynchronously 
        // (we'll return before this completes, relying on status updates)
        this.commandManager.executeCommand(command.id).catch(error => {
          console.error('Error during command execution:', error);
          // Error will be handled by the status listener
        });
        
        const assistantMessage = `Successfully completed: ${operationPlan.description.toLowerCase()}.`;
        
        return {
          processId: pid,
          assistantMessage: assistantMessage,
          command: command
        };
      } catch (error) {
        console.error('Error generating Excel operations:', error);
        
        sm.updateStatus(pid, {
          stage: ProcessStage.CommandPlanning,
          status: ProcessStatus.Error,
          message: `I failed to complete the operation. Please try again`
        });
        
        // If there was an error with the fallback and we're streaming, let the user know
        if (stream) {
          // Add a line break to separate from previous messages
          stream('\n\n');
          
          // Send the first recovery message immediately
          const randomMessage = recoveryMessages[Math.floor(Math.random() * recoveryMessages.length)];
          stream(randomMessage);
          
          // Set up a new interval for recovery messages
          let recoveryIndex = 0;
          const recoveryInterval = setInterval(() => {
            recoveryIndex++;
            if (recoveryIndex < recoveryMessages.length) {
              stream('\n\n' + recoveryMessages[recoveryIndex]);
            } else {
              // Stop after we've shown all recovery messages
              clearInterval(recoveryInterval);
            }
          }, 3000); // Show a new message every 3 seconds
          
          // Clear the recovery interval after 15 seconds to prevent endless messages
          setTimeout(() => {
            clearInterval(recoveryInterval);
          }, 15000);
        }
        return {
          processId: pid,
          assistantMessage: "I'm sorry, I encountered an error while planning the Excel operations. Please try again with a more specific request.",
          command: null
        };
      }
    }
}