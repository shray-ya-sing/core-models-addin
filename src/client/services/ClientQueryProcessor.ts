// src/client/services/ClientQueryProcessor.ts
import { v4 as uuidv4 } from 'uuid';
import {
  ClientAnthropicService,
  ModelType
} from './ClientAnthropicService';
import { ClientKnowledgeBaseService } from './ClientKnowledgeBaseService';
import { ClientWorkbookStateManager } from './ClientWorkbookStateManager';
import { ClientSpreadsheetCompressor } from './ClientSpreadsheetCompressor';
import { ClientCommandManager } from './ClientCommandManager';
import { ClientExcelOperationGenerator } from './ClientExcelOperationGenerator';
import { ExcelCommandPlan } from '../models/ExcelOperationModels';
import {
  Command,
  CommandStatus,
  CompressedWorkbook,
  QueryType as CommandModelQueryType,
  SheetState,
  WorkbookState,
  OperationType
} from '../models/CommandModels';
import {
  ProcessStage,
  ProcessStatus,
  ProcessStatusManager
} from '../models/ProcessStatusModels';
import { QueryContextBuilder } from './QueryContextBuilder';
import { ChunkLocatorService } from './ChunkLocatorService';
import { EmbeddingStore, SimilaritySearchResult } from './ChunkLocatorService'; // Temporary import, will be replaced with actual implementation





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
  private workbookManager: ClientWorkbookStateManager | null;
  private compressor: ClientSpreadsheetCompressor | null;
  private commandManager: ClientCommandManager | null;
  private currentQuerySteps: QueryStep[] = [];
  // Query context builder for more efficient state capture
  private queryContextBuilder: QueryContextBuilder;
  // Chunk locator service for identifying relevant chunks
  private chunkLocator: ChunkLocatorService | null = null;
  // Embedding store for similarity search
  private embeddingStore: EmbeddingStore | null = null;
  // Whether advanced chunk location is enabled
  private useAdvancedChunkLocation: boolean = true;

  constructor(params: {
    anthropic: ClientAnthropicService;
    kbService?: ClientKnowledgeBaseService | null;
    workbookStateManager?: ClientWorkbookStateManager | null;
    compressor?: ClientSpreadsheetCompressor | null;
    commandManager?: ClientCommandManager | null;
    useAdvancedChunkLocation?: boolean;
  }) {
    this.anthropic = params.anthropic;
    this.kbService = params.kbService ?? null;
    this.workbookManager = params.workbookStateManager ?? null;
    this.compressor = params.compressor ?? null;
    this.commandManager = params.commandManager ?? null;
    this.currentQuerySteps = [];
    this.useAdvancedChunkLocation = params.useAdvancedChunkLocation ?? true;
    
    // Create the query context builder
    this.queryContextBuilder = new QueryContextBuilder(
      this.workbookManager, 
      this.workbookManager.getMetadataCache(),
      this.chunkLocator);
    
    // Initialize advanced chunk location components if enabled
    if (this.useAdvancedChunkLocation && this.workbookManager) {
      this.initializeChunkLocator();
    }
  }
  
  /**
   * Initialize the chunk locator service
   */
  private async initializeChunkLocator(): Promise<void> {
    console.log('%c Initializing advanced chunk location components', 'background: #8e44ad; color: #ecf0f1; font-size: 12px; padding: 2px 5px;');
    
    try {
      // In a future implementation, we'd create a real EmbeddingStore
      // For now, we'll use our simple placeholder implementation
      this.embeddingStore = {} as EmbeddingStore; // Placeholder
      
      // Create the chunk locator service
      this.chunkLocator = new ChunkLocatorService({
        metadataCache: this.workbookManager.getMetadataCache(),
        embeddingStore: this.embeddingStore,
        dependencyAnalyzer: this.workbookManager.getDependencyAnalyzer(),
        anthropicService: this.anthropic,
        activeSheetName: this.workbookManager.getActiveSheetName()
      });
      
      // Attach the chunk locator to the query context builder
      this.queryContextBuilder.setChunkLocator(this.chunkLocator);
      
      console.log('%c Advanced chunk location components initialized successfully', 'color: #2ecc71');
    } catch (error) {
      console.error('Error initializing chunk locator:', error);
      console.log('%c Falling back to standard chunk identification', 'color: #e74c3c');
      this.useAdvancedChunkLocation = false;
    }
  }


  /* -----------------------  Public Entry Point  ----------------------- */

  async processQuery(
    userQuery: string,
    streamingCB?: (chunk: string) => void,
    chatHistory: Array<{role: string, content: string}> = []
  ): Promise<QueryProcessorResult> {
    const processId = uuidv4();
    const status = ProcessStatusManager.getInstance();
    
    console.group(`%c Query Processing (ID: ${processId.substring(0, 8)})`, 'background: #222; color: #f39c12; font-size: 14px; padding: 3px 6px;');
    console.log(`%c Received query: "${userQuery}"`, 'color: #2c3e50; font-size: 13px; font-weight: bold;');
    console.time('Query processing time');

    /* 1. Classify query with the LLM */
    const queryClassification = await this.classifyQueryAndDecompose(
      userQuery,
      processId,
      chatHistory
    );

    if (!queryClassification) {
      console.log(`%c No steps found in the query classification result`, 'background: #c0392b; color: #fff; font-size: 12px; padding: 2px 5px;');
      console.timeEnd('Query processing time');
      console.groupEnd();
      return {
        processId,
        assistantMessage: "I couldn't process your request. Please try again with a clearer instruction."
      };
    }

    else {
      // Process steps sequentially with proper async/await
      // Use for...of loop instead of forEach for proper async handling
      for (const step of queryClassification.steps) {
        const queryType = step.step_type;
        const query = step.step_specific_query;
        const stepIndex = step.step_index;
        
        // Log that you are completing this step
        console.log(`%c Completing step ${stepIndex} of type ${queryType} with query: ${query}`, 'color: #2ecc71');

        /* 4. Route execution */
        switch (queryType) {
          case QueryType.Greeting:
            // For greetings, no need to capture workbook state
            console.log(`%c Processing greeting: "${query}"`, `background: ${this.getClassificationColor(QueryType.Greeting)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
            console.timeEnd('Query processing time');
            console.groupEnd();
            return this.handleGreeting(query, streamingCB, processId);
              
          case QueryType.WorkbookQuestion:
          case QueryType.WorkbookQuestionWithKB:
          case QueryType.WorkbookCommand:
          case QueryType.WorkbookCommandWithKB:

            // Use our contextBuilder to get chunks via the query-dependent selection
            const queryContext = await this.queryContextBuilder.buildContextForQuery(
              this.convertToCommandModelQueryType(queryType as QueryType),
              chatHistory,
              query
            );
            
            // Convert the context to JSON format
            const workbookJSON = this.queryContextBuilder.contextToJson(queryContext);
            
            // Route to appropriate handler based on query type
            if (queryType === QueryType.WorkbookQuestion) {
              console.log(`%c Processing workbook question: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookQuestion)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
              console.timeEnd('Query processing time');
              console.groupEnd();
              return this.answerWorkbookQuestion(
                query,
                workbookJSON,
                streamingCB,
                processId,
                chatHistory
              );
            } else if (queryType === QueryType.WorkbookQuestionWithKB) {
              console.log(`%c Processing workbook question with KB: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookQuestionWithKB)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
              console.timeEnd('Query processing time');
              console.groupEnd();
              return this.answerWorkbookQuestion(
                query,
                workbookJSON,
                streamingCB,
                processId,
                chatHistory
              );
            } else if (queryType === QueryType.WorkbookCommand) {
              console.log(`%c Processing workbook command: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookCommand)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
              console.timeEnd('Query processing time');
              console.groupEnd();
              return this.runWorkbookCommand(
                query,
                workbookJSON,
                processId,
                chatHistory
              );
            } else if (queryType === QueryType.WorkbookCommandWithKB) {
              console.log(`%c Processing workbook command with KB: "${query}"`, `background: ${this.getClassificationColor(QueryType.WorkbookCommandWithKB)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
              console.timeEnd('Query processing time');
              console.groupEnd();
              return this.runWorkbookCommand(
                query,
                workbookJSON,
                processId,
                chatHistory
              );
            }
            else if (queryType === QueryType.Unknown) {
              console.log(`%c Unable to classify query: "${query}"`, `background: ${this.getClassificationColor(QueryType.Unknown)}; color: #fff; font-size: 12px; padding: 2px 5px;`);
              console.timeEnd('Query processing time');
              console.groupEnd();
              return {
                processId,
                assistantMessage:
                  "I'm sorry, I'm not sure how to handle that request yet."
              };
            }        
        }
      }
      
      // After processing all steps, return a success result
      console.log(`%c All steps processed successfully`, 'background: #27ae60; color: #fff; font-size: 12px; padding: 2px 5px;');
      console.timeEnd('Query processing time');
      console.groupEnd();
      return {
        processId,
        assistantMessage: "I completed all steps of your request"
      };
    }
  }

  /* --------------------------- Helper Methods --------------------------- */

 
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

  private async classifyQueryAndDecompose(
    query: string,
    _pid: string, 
    chatHistory: Array<{role: string, content: string}> = []
  ): Promise<QueryClassification> {
    try {
      const res = await this.anthropic.classifyQueryAndDecompose(query, chatHistory);
      
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

  /**
   * Helper method to get a color for each query type for styled console logs
   * @param queryType The query type to get a color for
   * @returns A CSS color string
   */
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
      pid: string
    ): Promise<QueryProcessorResult> {
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.QueryProcessing,
        status: ProcessStatus.Pending,
        message: 'Generating greeting…'
      });
  
      const resp = await this.anthropic.generateChatResponse(query, stream);
  
      sm.updateStatus(pid, {
        stage: ProcessStage.QueryProcessing,
        status: ProcessStatus.Success,
        message: 'Greeting sent.'
      });
  
      return { processId: pid, assistantMessage: resp.assistantMessage };
    }
  
  
    private async answerWorkbookQuestion(
      query: string,
      workbookJSON: string,
      stream: ((c: string) => void) | undefined,
      pid: string,
      chatHistory: Array<{role: string, content: string}>
    ): Promise<QueryProcessorResult> {
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.ResponseGeneration,
        status: ProcessStatus.Pending,
        message: 'Answering workbook question…'
      });
  
      const resp = await this.anthropic.generateWorkbookExplanation(
        query,
        workbookJSON,
        stream,
        chatHistory
      );
  
      sm.updateStatus(pid, {
        stage: ProcessStage.ResponseGeneration,
        status: ProcessStatus.Success,
        message: 'Answered question.'
      });
  
      return {
        processId: pid,
        assistantMessage: resp.assistantMessage
      };
    }
  
    private async answerWorkbookQuestionWithKB(
      query: string,
      workbookJSON: string,
      stream: ((c: string) => void) | undefined,
      pid: string,
      chatHistory: Array<{role: string, content: string}>
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
  
        return this.answerWorkbookQuestion(query, context, stream, pid, chatHistory);
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
      pid: string,
      chatHistory: Array<{ role: string; content: string }>
    ): Promise<QueryProcessorResult> {
      /* 1. Initialize status tracking */
      const sm = ProcessStatusManager.getInstance();
      sm.updateStatus(pid, {
        stage: ProcessStage.CommandPlanning,
        status: ProcessStatus.Pending,
        message: 'Command plan ready...'
      });
    
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
        const operationPlan = await operationGenerator.generateOperations(
          query,
          workbookJSON,
          chatHistory
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
        
        // Generate a simpler user-friendly message without mentioning operation count
        const assistantMessage = `I'll ${operationPlan.description.toLowerCase()}.`;
        
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
        
        return {
          processId: pid,
          assistantMessage: "I'm sorry, I encountered an error while planning the Excel operations. Please try again with a more specific request.",
          command: null
        };
      }
    }
}