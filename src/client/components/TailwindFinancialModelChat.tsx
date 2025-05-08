import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import { v4 as uuidv4 } from 'uuid';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import VersionHistoryView from './VersionHistoryView';
import { ClientCommandManager } from '../services/ClientCommandManager';
import { ClientCommandExecutor } from '../services/ClientCommandExecutor';
import { ClientWorkbookStateManager } from '../services/ClientWorkbookStateManager';
import { ClientSpreadsheetCompressor } from '../services/ClientSpreadsheetCompressor';
import { ClientExcelCommandAdapter } from '../services/ClientExcelCommandAdapter';
import { ClientExcelCommandInterpreter } from '../services/ClientExcelCommandInterpreter';
import { ClientAnthropicService } from '../services/ClientAnthropicService';
import { ClientKnowledgeBaseService } from '../services/ClientKnowledgeBaseService';
import { ClientQueryProcessor, QueryProcessorResult } from '../services/ClientQueryProcessor';
import { VersionHistoryProvider } from '../services/versioning/VersionHistoryProvider';
import { VersionEventType } from '../models/VersionModels';
import { Command, CommandStatus } from '../models/CommandModels';
import { ProcessStatusManager, ProcessStatus, ProcessStage } from '../models/ProcessStatusModels';
import { TypewriterEffect } from './TypewriterEffect';
import StatusIndicator, { StatusType } from './StatusIndicator';
import ProcessStatusTracker from './ProcessStatusTracker';
import { getStageName } from './ProcessStatusTracker';
import config from '../config';
import { AIApprovalSystem } from '../services/AIApprovalSystem';
import { PendingChangesTracker, PendingChange } from '../services/PendingChangesTracker';
import { ShapeEventHandler } from '../services/ShapeEventHandler';
import { SendIcon, FileIcon, CheckIcon, AlertCircleIcon } from './icons';
import PendingChangesBar from './PendingChangesBar';

// Message type definition
interface ChatMessage {
  role: 'user' | 'assistant' | 'system' | 'status';
  content: string;
  isStreaming?: boolean;
  status?: StatusType;
  stage?: string;
}

// Conversation session interface
interface ConversationSession {
  id: string;
  title: string;
  messages: ChatMessage[];
  lastUpdated: number; // timestamp
  createdAt: number; // timestamp
  workbookId?: string; // Identifier for the associated workbook
}

interface TailwindFinancialModelChatProps {
  newConversationTrigger?: number;
  showPastConversationsTrigger?: number;
  showVersionHistoryTrigger?: number;
}

const TailwindFinancialModelChat: React.FC<TailwindFinancialModelChatProps> = ({ 
  newConversationTrigger = 0,
  showPastConversationsTrigger = 0,
  showVersionHistoryTrigger = 0
}) => {
  // Storage key for conversations in localStorage
  const STORAGE_KEY = 'excel-addin-conversations';
  
  // Current conversation and messages state
  const [currentSession, setCurrentSession] = useState<string>(''); // Current session ID
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [userInput, setUserInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [hasUserSentMessage, setHasUserSentMessage] = useState(false);
  const [servicesReady, setServicesReady] = useState(false);
  const [streamingResponse, setStreamingResponse] = useState<string>('');
  const [isStreaming, setIsStreaming] = useState<boolean>(false);
  
  // All saved conversations
  const [sessions, setSessions] = useState<ConversationSession[]>([]);
  
  // Current workbook ID
  const [currentWorkbookId, setCurrentWorkbookId] = useState<string>('');
  
  // Filter sessions by current workbook ID
  const getWorkbookSessions = (allSessions: ConversationSession[]): ConversationSession[] => {
    // If we don't have a workbook ID yet, return all sessions
    if (!currentWorkbookId) {
      return allSessions;
    }
    
    // Filter sessions to only include those for the current workbook
    // Also include sessions without a workbookId (for backward compatibility)
    return allSessions.filter(session => 
      !session.workbookId || session.workbookId === currentWorkbookId
    );
  };
  
  // State to control showing all conversations
  const [showAllConversations, setShowAllConversations] = useState(false);
  
  // State to control showing the past conversations view
  const [showPastConversationsView, setShowPastConversationsView] = useState(false);
  
  // State to control showing the version history view
  const [showVersionHistoryView, setShowVersionHistoryView] = useState(false);
  
  // Refs
  const chatContainerRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Service instances
  const [commandManager, setCommandManager] = useState<ClientCommandManager | null>(null);
  const [commandExecutor, setCommandExecutor] = useState<ClientCommandExecutor | null>(null);
  const [workbookStateManager, setWorkbookStateManager] = useState<ClientWorkbookStateManager | null>(null);
  const [spreadsheetCompressor, setSpreadsheetCompressor] = useState<ClientSpreadsheetCompressor | null>(null);
  const [anthropicService, setAnthropicService] = useState<ClientAnthropicService | null>(null);
  const [knowledgeBaseService, setKnowledgeBaseService] = useState<ClientKnowledgeBaseService | null>(null);
  const [queryProcessor, setQueryProcessor] = useState<ClientQueryProcessor | null>(null);
  const [commandInterpreter, setCommandInterpreter] = useState<ClientExcelCommandInterpreter | null>(null);
  const versionHistoryProviderRef = useRef<VersionHistoryProvider>(new VersionHistoryProvider());
  const [pendingChangesTracker, setPendingChangesTracker] = useState<PendingChangesTracker | null>(null);
  const [shapeEventHandler, setShapeEventHandler] = useState<ShapeEventHandler | null>(null);
  const [approvalEnabled, setApprovalEnabled] = useState<boolean>(false);

  // Status trackers
  const [statusManager] = useState<ProcessStatusManager>(() => ProcessStatusManager.getInstance());
  const [currentStatus, setCurrentStatus] = useState<ProcessStatus | null>(null);
  
  // Pending changes state
  const [pendingChanges, setPendingChanges] = useState<PendingChange[]>([]);
  
  // Function to refresh pending changes
  const refreshPendingChanges = useCallback(() => {
    if (pendingChangesTracker && currentWorkbookId) {
      const changes = pendingChangesTracker.getPendingChanges(currentWorkbookId);
      setPendingChanges(changes);
    }
  }, [pendingChangesTracker, currentWorkbookId]);
  
  // Functions to handle accept/reject actions
  const handleAcceptAll = useCallback(async () => {
    if (pendingChangesTracker && pendingChanges.length > 0) {
      for (const change of pendingChanges) {
        await pendingChangesTracker.acceptChange(change.id);
      }
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, pendingChanges, refreshPendingChanges]);
  
  const handleRejectAll = useCallback(async () => {
    if (pendingChangesTracker && pendingChanges.length > 0) {
      for (const change of pendingChanges) {
        await pendingChangesTracker.rejectChange(change.id);
      }
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, pendingChanges, refreshPendingChanges]);
  
  const handleAcceptChange = useCallback(async (changeId: string) => {
    if (pendingChangesTracker) {
      await pendingChangesTracker.acceptChange(changeId);
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, refreshPendingChanges]);
  
  const handleRejectChange = useCallback(async (changeId: string) => {
    if (pendingChangesTracker) {
      await pendingChangesTracker.rejectChange(changeId);
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, refreshPendingChanges]);

  // Get current workbook ID
  const getCurrentWorkbookId = async (): Promise<string> => {
    try {
      // Check if Office and Excel are available
      if (typeof Excel === 'undefined') {
        console.warn('Excel API not available');
        return 'unknown-workbook';
      }
      
      return await Excel.run(async (context) => {
        // Get the workbook properties
        const workbook = context.workbook;
        workbook.load('name');
        
        await context.sync();
        
        // Use the workbook name as the ID, or a fallback if not available
        const workbookId = workbook.name || `workbook-${new Date().getTime()}`;
        return workbookId;
      });
    } catch (error) {
      console.error('Error getting workbook ID:', error);
      return 'unknown-workbook';
    }
  };
  
  // Initialize services
  // Effect to periodically refresh pending changes
  useEffect(() => {
    if (pendingChangesTracker && currentWorkbookId) {
      // Initial refresh
      refreshPendingChanges();
      
      // Set up interval to refresh pending changes
      const intervalId = window.setInterval(() => {
        refreshPendingChanges();
      }, 2000); // Refresh every 2 seconds
      
      return () => {
        window.clearInterval(intervalId);
      };
    }
    
    // Return empty cleanup function if conditions aren't met
    return () => {};
  }, [pendingChangesTracker, currentWorkbookId, refreshPendingChanges]);

  useEffect(() => {
    // Initialize all client services
    const initializeClientServices = async () => {
      // Get the current workbook ID
      const workbookId = await getCurrentWorkbookId();
      setCurrentWorkbookId(workbookId);
      console.log('Current workbook ID:', workbookId);
      try {
        // Log initialization
        console.log('%c Initializing client services...', 'background: #222; color: #bada55; font-size: 14px');
        
        // Check for API key
        if (!config.anthropicApiKey) {
          console.warn('No Anthropic API key found in configuration');
        }
        
        // Create service instances
        // Create a single interpreter instance that will be shared by all components
        const interpreter = new ClientExcelCommandInterpreter();
        
        // Initialize version history system first
        console.log(`🔄 [TailwindFinancialModelChat] Initializing version history system with workbookId: ${workbookId}`);
        versionHistoryProviderRef.current.initialize(interpreter);
        versionHistoryProviderRef.current.setCurrentWorkbookId(workbookId);
        
        // Verify the version history system is properly initialized
        console.log(`✅ [TailwindFinancialModelChat] Version history system initialized with interpreter:`, {
          hasActionRecorder: !!interpreter.getActionRecorder(),
          workbookId: interpreter.getCurrentWorkbookId() || 'not set'
        });
        
        // Now create the other services
        const stateManager = new ClientWorkbookStateManager();
        const compressor = new ClientSpreadsheetCompressor();
        const executor = new ClientCommandExecutor(stateManager);
        
        // Create the command adapter with our shared interpreter instance
        const adapter = new ClientExcelCommandAdapter(interpreter);
        
        // Create the command manager with our shared adapter instance
        const manager = new ClientCommandManager(executor, stateManager, adapter);
        
        // Log the shared components to verify they're properly connected
        console.log(`🔄 [TailwindFinancialModelChat] Shared components:`, {
          interpreter: interpreter,
          adapter: adapter,
          manager: manager
        });
        const anthropic = new ClientAnthropicService(config.anthropicApiKey, config.openaiApiKey);
        const knowledgeBase = new ClientKnowledgeBaseService(config.knowledgeBaseApiUrl);
        const processor = new ClientQueryProcessor({
          anthropic,
          kbService: knowledgeBase,
          workbookStateManager: stateManager,
          compressor,
          commandManager: manager
        });
        
        // Set up event listeners for workbook changes
        await stateManager.setupChangeListeners();
        
        // Initialize AI Approval System
        console.log(`🔄 [TailwindFinancialModelChat] Initializing AI approval system...`);
        const versionHistoryService = versionHistoryProviderRef.current.getVersionHistoryService();
        const { pendingChangesTracker: pct, shapeEventHandler: seh } = AIApprovalSystem.initialize(interpreter, versionHistoryService);
        
        // Set the current workbook ID on the shape event handler
        seh.setCurrentWorkbookId(workbookId);
        
        // Start the shape event handler polling
        seh.startPolling();
        
        // Update state with service instances
        setWorkbookStateManager(stateManager);
        setSpreadsheetCompressor(compressor);
        setCommandExecutor(executor);
        setCommandManager(manager);
        setAnthropicService(anthropic);
        setKnowledgeBaseService(knowledgeBase);
        setQueryProcessor(processor);
        setCommandInterpreter(interpreter);
        setPendingChangesTracker(pct);
        setShapeEventHandler(seh);
        
        // Enable approval workflow by default
        interpreter.setRequireApproval(true);
        setApprovalEnabled(true);
        console.log(`✅ [TailwindFinancialModelChat] AI approval system initialized and enabled`);
        
        // Register for command updates
        const unsubscribeCommandUpdate = manager.onCommandUpdate((command) => {
          console.log('Command update received:', command);
          
          // If the command is completed, log it but don't add a message to the chat
          if (command.status === CommandStatus.Completed) {
            console.log(`Command "${command.description}" completed successfully.`);
            // Removed system message
          } else if (command.status === CommandStatus.Failed) {
            console.error(`Command "${command.description}" failed: ${command.error || 'Unknown error'}`);
            // Only show error messages, not success messages
            setMessages(prev => [...prev, { 
              role: 'system', 
              content: `Command "${command.description}" failed: ${command.error || 'Unknown error'}` 
            }]);
          }
        });
                
        // Set up process status manager listener
        const processManager = ProcessStatusManager.getInstance();
        const unsubscribeProcessEvents = processManager.addListener((event) => {
          console.log('Process Status Event:', event);
          
          // Map process status to StatusIndicator status
          let statusType = StatusType.Idle;
          switch (event.status) {
            case ProcessStatus.Pending:
              statusType = StatusType.Pending;
              break;
            case ProcessStatus.Success:
              statusType = StatusType.Success;
              break;
            case ProcessStatus.Error:
              statusType = StatusType.Error;
              break;
          }
          
          // Add process status as a message in the chat
          setMessages(prev => {
            // Find the last user message index
            const lastUserMsgIndex = [...prev].reverse().findIndex(m => m.role === 'user');
            
            // If we found a user message
            if (lastUserMsgIndex >= 0) {
              const actualIndex = prev.length - 1 - lastUserMsgIndex;
              
              // First, remove any pending messages for this stage if we're getting a success/error
              let newMessages = [...prev];
              if (statusType === StatusType.Success || statusType === StatusType.Error) {
                newMessages = newMessages.filter(m => !(m.role === 'status' && m.stage === event.stage && m.status === StatusType.Pending));
              }
              
              // Now find if we have any existing status message for this stage after filtering
              const existingStatusIndex = newMessages.findIndex(m => m.role === 'status' && m.stage === event.stage && m.status === statusType);
              
              // Create the status message
              const statusMessage = {
                role: 'status' as const,
                content: event.message,
                status: statusType,
                stage: event.stage
              };
              
              if (existingStatusIndex >= 0) {
                // Update existing message with same status
                newMessages[existingStatusIndex] = statusMessage;
              } else {
                // Find the right position to insert the new message
                // It should go after the user message and any existing status messages for this query
                let insertIndex = actualIndex + 1;
                
                // Find the last status message related to the current user message
                for (let i = insertIndex; i < newMessages.length; i++) {
                  if (newMessages[i].role === 'status') {
                    insertIndex = i + 1;
                  } else if (newMessages[i].role === 'assistant') {
                    // Stop at the assistant message
                    break;
                  }
                }
                
                // Insert at the determined position
                newMessages.splice(insertIndex, 0, statusMessage);
              }
              
              return newMessages;
            }
            
            return prev;
          });
        });
        
        // Mark services as ready
        console.log('%c All services initialized successfully!', 'background: #222; color: #2ecc71; font-size: 14px');
        setServicesReady(true);
        
        // Return cleanup function
        return () => {
          if (unsubscribeCommandUpdate) unsubscribeCommandUpdate();
          if (unsubscribeProcessEvents) unsubscribeProcessEvents();
          
          // Stop the shape event handler polling
          if (seh) {
            seh.stopPolling();
            console.log('🛑 [TailwindFinancialModelChat] Stopped shape event handler polling');
          }
        };
      } catch (error) {
        console.error('Error initializing client services:', error);
        setMessages(prev => [...prev, { 
          role: 'system', 
          content: `Error initializing services: ${error.message}` 
        }]);
        return () => {}; // Return empty cleanup function for error case
      }
    };
    
    const cleanup = initializeClientServices();
    
    // Cleanup function
    return () => {
      if (cleanup) {
        cleanup.then(cleanupFn => {
          if (cleanupFn) cleanupFn();
          return;
        }).catch(err => {
          console.error('Error in cleanup function:', err);
        });
      }
    };
  }, []);

  // Load saved conversations from localStorage on component mount
  useEffect(() => {
    const loadSavedSessions = () => {
      try {
        const savedSessions = localStorage.getItem(STORAGE_KEY);
        if (savedSessions) {
          const parsedSessions = JSON.parse(savedSessions) as ConversationSession[];
          // Sort sessions by last updated timestamp (newest first)
          const sortedSessions = parsedSessions.sort((a, b) => b.lastUpdated - a.lastUpdated);
          setSessions(sortedSessions);
        }
      } catch (error) {
        console.error('Error loading saved sessions:', error);
      }
    };
    
    loadSavedSessions();
  }, []);
  
  // Set current session based on workbook ID when it changes
  useEffect(() => {
    if (!currentWorkbookId || sessions.length === 0) return;
    
    console.log('Setting session based on workbook ID:', currentWorkbookId);
    
    // Get workbook-specific sessions
    const workbookSessions = getWorkbookSessions(sessions);
    
    // If there are workbook-specific sessions, set the most recent one as current
    if (workbookSessions.length > 0) {
      setCurrentSession(workbookSessions[0].id);
      setMessages(workbookSessions[0].messages);
      console.log('Loaded existing session for workbook:', workbookSessions[0].title);
    } else {
      // If no workbook-specific sessions, create a new one
      console.log('No sessions for current workbook, creating new session');
      createNewSession();
    }
  }, [currentWorkbookId, sessions.length]);

  // Auto-scroll to bottom when messages change
  useEffect(() => {
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);
  
  // Update localStorage when sessions change
  useEffect(() => {
    if (sessions.length > 0) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    }
  }, [sessions]);
  
  // Update current session's messages when messages state changes
  useEffect(() => {
    if (currentSession && messages.length > 0) {
      updateSessionMessages(currentSession, messages);
    }
  }, [messages]);
  
  // Session management functions
  const createNewSession = () => {
    const newSessionId = uuidv4();
    const newSession: ConversationSession = {
      id: newSessionId,
      title: 'New Conversation', // Default title
      messages: [],
      lastUpdated: Date.now(),
      createdAt: Date.now(),
      workbookId: currentWorkbookId // Associate with current workbook
    };
    
    setSessions(prev => [newSession, ...prev]);
    setCurrentSession(newSessionId);
    setMessages([]);
    setHasUserSentMessage(false);
    return newSessionId;
  };
  
  const getSessionTitle = (messages: ChatMessage[]) => {
    // Find first user message to use as title
    const firstUserMessage = messages.find(msg => msg.role === 'user');
    if (firstUserMessage) {
      // Truncate to reasonable title length (max 30 chars)
      const truncatedContent = firstUserMessage.content.substring(0, 30);
      return truncatedContent + (firstUserMessage.content.length > 30 ? '...' : '');
    }
    return 'New Conversation';
  };
  
  const updateSessionMessages = (sessionId: string, updatedMessages: ChatMessage[]) => {
    setSessions(prev => {
      return prev.map(session => {
        if (session.id === sessionId) {
          // Update the title if it's still the default and we have user messages
          const title = session.title === 'New Conversation' 
            ? getSessionTitle(updatedMessages)
            : session.title;
            
          return {
            ...session,
            messages: updatedMessages,
            title,
            lastUpdated: Date.now()
          };
        }
        return session;
      });
    });
  };
  
  const loadSession = (sessionId: string) => {
    const session = sessions.find(s => s.id === sessionId);
    if (session) {
      setCurrentSession(sessionId);
      setMessages(session.messages);
      setHasUserSentMessage(session.messages.length > 0);
    }
  };
  
  const formatTimeAgo = (timestamp: number) => {
    const now = Date.now();
    const secondsAgo = Math.floor((now - timestamp) / 1000);
    
    if (secondsAgo < 60) return `${secondsAgo}s`;
    if (secondsAgo < 3600) return `${Math.floor(secondsAgo / 60)}m`;
    if (secondsAgo < 86400) return `${Math.floor(secondsAgo / 3600)}h`;
    return `${Math.floor(secondsAgo / 86400)}d`;
  };

  // Load the most recent session when first opened (but don't create a new one)
  useEffect(() => {
    if (!currentSession && sessions.length > 0) {
      // If there are existing sessions but none selected, select the most recent one
      setCurrentSession(sessions[0].id);
      setMessages(sessions[0].messages);
      setHasUserSentMessage(sessions[0].messages.length > 0);
    }
  }, [sessions]);
  
  // Handle sending a message
  const handleSendMessage = async () => {
      console.log('%c handleSendMessage called', 'background: #222; color: #3498db; font-size: 14px');
      console.log('servicesReady:', servicesReady, 'queryProcessor:', !!queryProcessor);
      
      if (!servicesReady || !queryProcessor) {
        console.warn('Services not ready yet!');
        setMessages(prev => [...prev, {
          role: 'system',
          content: 'Services are still initializing, please wait a moment...'
        }]);
        return;
      }
      
      if (!userInput.trim()) return;
      
      // Add user message to chat
      // If no current session, create one
      if (!currentSession) {
        const newSessionId = createNewSession();
        setCurrentSession(newSessionId);
      }
      
      setMessages(prev => [...prev, { role: 'user', content: userInput }]);
      setUserInput('');
      setIsLoading(true);
      setHasUserSentMessage(true);
      
      try {
        console.log('%c Sending query to processor:', 'background: #222; color: #f39c12; font-size: 14px', userInput.substring(0, 50));
        
        // Create a streaming response handler
        const handleStreamingResponse = (chunk: string) => {
          console.log('Received chunk:', chunk.substring(0, 20) + (chunk.length > 20 ? '...' : ''));
          setStreamingResponse(prev => prev + chunk); // Append each chunk to our streaming state
          
          // Also update the messages array with the current partial response
          setMessages(prev => {
            // Create a new array to avoid mutating state directly
            const newMessages = [...prev];
            
            // Find the assistant message if it exists (replace it) or add a new one
            const assistantMsgIndex = newMessages.findIndex(m => 
              m.role === 'assistant' && m.isStreaming === true
            );
            
            if (assistantMsgIndex >= 0) {
              // Update existing message
              newMessages[assistantMsgIndex] = {
                ...newMessages[assistantMsgIndex],
                content: prev[assistantMsgIndex].content + chunk
              };
            } else {
              // Add new message
              newMessages.push({
                role: 'assistant',
                content: chunk,
                isStreaming: true
              } as ChatMessage);
            }
            
            return newMessages;
          });
        };
        
        // Set streaming mode on
        setIsStreaming(true);
        setStreamingResponse('');
        
        // Create a safe wrapper that doesn't throw
        const safeProcessQuery = async () => {
          try {
            // Convert our messages to the format expected by the ClientQueryProcessor
            const chatHistory = messages
              // Only include user, assistant, and system messages (exclude status messages)
              .filter(msg => msg.role === 'user' || msg.role === 'assistant' || msg.role === 'system')
              // Exclude streaming messages
              .filter(msg => !msg.isStreaming)
              // Only include the last 10 messages to avoid token limits
              .slice(-10)
              .map(msg => ({
                // Explicitly cast to the allowed roles to satisfy TypeScript
                role: msg.role as 'user' | 'assistant' | 'system',
                content: msg.content
              }));
            
            console.log(`%c Passing ${chatHistory.length} messages as conversation history`, 'color: #3498db');
            
            // Pass the streaming handler and chat history to the query processor
            return await queryProcessor.processQuery(userInput, handleStreamingResponse, chatHistory);
          } catch (innerError) {
            console.error('Error in queryProcessor.processQuery:', innerError);
            setIsStreaming(false);
            return { 
              processId: uuidv4(),
              assistantMessage: "I encountered an error processing your request. This might be due to API connection issues. Please try again later.",
              command: null
            } as QueryProcessorResult;
          }
        };
        
        // Process the query using our new query processor with safe wrapper
        const result = await safeProcessQuery();
        
        // Streaming is complete
        setIsStreaming(false);
        
        // If we weren't streaming (or streaming failed), add the complete assistant response
        // Otherwise, the streaming handler would have already added the message
        if (!isStreaming) {
          setMessages(prev => {
            // Check if we already have a streaming message
            const streamingIndex = prev.findIndex(m => m.role === 'assistant' && m.isStreaming === true);
            
            if (streamingIndex >= 0) {
              // Replace the streaming message with the final one
              const newMessages = [...prev];
              newMessages[streamingIndex] = { 
                role: 'assistant', 
                content: result.assistantMessage 
              };
              return newMessages;
            } else {
              // No streaming message exists, add a new one
              return [...prev, { 
                role: 'assistant', 
                content: result.assistantMessage 
              }];
            }
          });
        } else {
          // Finalize any streaming message by removing the isStreaming flag
          setMessages(prev => prev.map(msg => 
            msg.isStreaming ? { ...msg, isStreaming: undefined } : msg
          ));
        }
        
        // If there's a command in the response, log it but don't show a message
        if (result.command) {
          console.log(`🔄 [TailwindFinancialModelChat] Command received: ${result.command.description}`);
          
          // Log the version history state
          if (commandInterpreter) {
            console.log(`🔄 [TailwindFinancialModelChat] Version history state:`, {
              hasActionRecorder: !!commandInterpreter.getActionRecorder(),
              workbookId: commandInterpreter.getCurrentWorkbookId() || 'not set'
            });
          }
          
          // Check if command is already being executed
          if (commandManager) {
            const command = commandManager.getCommand(result.command.id);
            if (command) {
              // Command exists, check its status
              if (command.status === 'running' || command.status === 'pending') {
                console.log(`🔔 [TailwindFinancialModelChat] Command ${result.command.id} is already being executed (status: ${command.status}). Skipping duplicate execution.`);
                // We'll just wait for the status updates via the listener
              } else if (command.status === 'completed') {
                console.log(`✅ [TailwindFinancialModelChat] Command ${result.command.id} is already completed. No need to execute again.`);
              } else if (command.status === 'failed') {
                console.log(`⚠️ [TailwindFinancialModelChat] Command ${result.command.id} previously failed. Not re-executing.`);
              } else {
                // Command exists but is in an unknown state, execute it
                console.log(`🔄 [TailwindFinancialModelChat] Executing command with ID: ${result.command.id}`);
                await commandManager.executeCommand(result.command.id);
                console.log(`✅ [TailwindFinancialModelChat] Command execution completed`);
              }
            } else {
              // Command doesn't exist in the manager yet, which is unusual
              // This might happen if there's a race condition where the command hasn't been added yet
              console.log(`❓ [TailwindFinancialModelChat] Command ${result.command.id} not found in command manager. This is unexpected.`);
            }
          }
        }
      } catch (error) {
        console.error('%c CRITICAL ERROR in message handling:', 'background: #f00; color: #fff; font-size: 14px', error);
        console.trace('Stack trace:');
        setMessages(prev => [...prev, { 
          role: 'system', 
          content: `An unexpected error occurred while processing your message. Technical details: ${error.message || 'Unknown error'}` 
        }]);
      } finally {
        setIsLoading(false);
      }
    };

    // Handle input key press (Enter to send)
      const handleKeyPress = (e: React.KeyboardEvent) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          handleSendMessage();
        }
      };
    
      // Handle input change
      const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
        setUserInput(e.target.value);
      };

  // Watch for new conversation triggers from the header component
  useEffect(() => {
    if (newConversationTrigger > 0) {
      createNewSession();
    }
  }, [newConversationTrigger]);
  
  // Watch for show past conversations triggers from the header component
  useEffect(() => {
    if (showPastConversationsTrigger > 0) {
      console.log('Show past conversations trigger received:', showPastConversationsTrigger);
      // Toggle the past conversations view instead of showing all conversations in-place
      setShowPastConversationsView(true);
      // Hide version history view if it's open
      setShowVersionHistoryView(false);
    }
  }, [showPastConversationsTrigger]);
  
  // Watch for show version history triggers from the header component
  useEffect(() => {
    if (showVersionHistoryTrigger > 0) {
      console.log('Show version history trigger received:', showVersionHistoryTrigger);
      
      // Ensure the version history provider has the current workbook ID
      if (commandInterpreter && currentWorkbookId) {
        console.log(`🔄 [TailwindFinancialModelChat] Ensuring version history is initialized before showing panel`);
        versionHistoryProviderRef.current.initialize(commandInterpreter);
        versionHistoryProviderRef.current.setCurrentWorkbookId(currentWorkbookId);
      }
      
      // Toggle the version history view
      setShowVersionHistoryView(true);
      // Hide past conversations view if it's open
      setShowPastConversationsView(false);
    }
  }, [showVersionHistoryTrigger, commandInterpreter, currentWorkbookId]);

  // Function to go back from past conversations view
  const closePastConversationsView = () => {
    setShowPastConversationsView(false);
  };
  
  // Function to go back from version history view
  const closeVersionHistoryView = () => {
    setShowVersionHistoryView(false);
  };

  return (
    <div className="flex flex-col h-full font-mono text-sm relative">
      {showVersionHistoryView ? (
        // Version History View - using the separate component
        <VersionHistoryView 
          onClose={closeVersionHistoryView} 
          workbookId={currentWorkbookId}
          versionHistoryProvider={versionHistoryProviderRef.current}
        />
      ) : showPastConversationsView ? (
        // Past Conversations View - full page transition
        <div className="flex flex-col h-full p-4 animate-fadeIn transition-all duration-300">
          {/* Header with back button */}
          <div className="flex justify-between items-center mb-4">
            <div className="flex items-center">
              <button 
                onClick={closePastConversationsView}
                className="text-gray-400 hover:text-white p-1 mr-2 rounded transition-colors"
                aria-label="Back"
              >
                <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 19l-7-7 7-7"></path>
                </svg>
              </button>
              <h2 className="font-medium text-white" style={{ 
                fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                fontSize: '12px'
              }}>Past Conversations</h2>
            </div>
          </div>
          
          {/* Scrollable Container for Past Conversations */}
          <div className="flex-1 overflow-y-auto">
            <div className="flex flex-col space-y-0.5">
              {sessions.length === 0 ? (
                <div className="text-gray-500 text-center py-2" style={{ 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                  fontSize: '11px' 
                }}>
                  No conversations yet
                </div>
              ) : (
                getWorkbookSessions(sessions).map(session => (
                  <div 
                    key={session.id} 
                    className={`flex justify-between items-center py-2 px-3 ${currentSession === session.id ? 'bg-gray-800/40' : 'hover:bg-black/20'} rounded-lg cursor-pointer border-l-2 border-black/40`}
                    onClick={() => {
                      loadSession(session.id);
                      closePastConversationsView();
                    }}
                  >
                    <div className="flex-grow text-white/80" style={{ 
                      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                      fontSize: '12px' 
                    }}>{session.title}</div>
                    <div className="ml-2" style={{ 
                      fontSize: '10px', 
                      opacity: 0.4, 
                      color: 'white', 
                      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                    }}>{formatTimeAgo(session.lastUpdated)}</div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      ) : !hasUserSentMessage ? (
        // Welcome screen with Tailwind CSS styling (similar to cascade-chat)
        <div className="flex flex-col h-full p-4">
          {/* Flex container that takes up all available space but allows Past Conversations to be at bottom */}
          <div className="flex-1 flex flex-col justify-center">
            {/* Main UI container - positioned to be just above Past Conversations */}
            <div className="flex flex-col items-start mb-8 w-full transition-all duration-300">
              {/* Logo without circular highlight - responsive sizing with direct inline styles */}
              <div className="mb-6">
                <img 
                  src="assets/cori-logo.svg" 
                  style={{ 
                    width: '42px', 
                    height: '42px', 
                    opacity: 0.9,
                    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
                  }} 
                  alt="Cori Logo" 
                />  
              </div>
              
              {/* Description with fixed font size */}
              <p 
                className="text-white/70 leading-relaxed text-left mb-6 max-w-sm" 
                style={{ 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                  fontSize: '12px'
                }}
              >
                Build financial models or analyze existing financial data with expert assistance
              </p>
              
              {/* Input area with separate send button - fixed sizes */}
              <div className="flex items-center w-full max-w-lg space-x-2">
                <textarea
                  ref={textareaRef}
                  className="flex-grow bg-transparent border border-gray-700 hover:border-gray-500 focus:border-blue-500 focus:ring-0 focus:outline-none p-2 text-white rounded-md placeholder-white/50 min-h-[36px] resize-none"
                  placeholder="Ask anything"
                  value={userInput}
                  onChange={handleInputChange}
                  onKeyDown={handleKeyPress}
                  disabled={isLoading || !servicesReady}
                  rows={1}
                  style={{ 
                    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                    fontSize: '12px'
                  }}
                />
                <button
                  className="text-gray-500 disabled:text-gray-700 hover:text-white hover:bg-gray-700/30 p-1.5 rounded-md transition-colors flex items-center justify-center"
                  onClick={handleSendMessage}
                  disabled={!userInput.trim() || isLoading || !servicesReady}
                  aria-label="Send message"
                  type="button"
                >
                  {/* Enhanced clickable right arrow */}
                  <svg style={{ width: '16px', height: '16px' }} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M5 12H19M19 12L13 6M19 12L13 18" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </button>
              </div>
            </div>
          </div>
          
          {/* Past conversations - positioned at bottom with fixed sizes */}
          <div className="w-full max-w-xl self-center past-conversations-section">
            <div className="mb-3">
              <h2 className="font-medium text-white" style={{ 
                fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                fontSize: '12px'
              }}>Past Conversations</h2>
            </div>
            <div className="flex flex-col space-y-0.5">
              {sessions.length === 0 ? (
                <div className="text-gray-500 text-center py-2" style={{ 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                  fontSize: '11px' 
                }}>
                  No conversations yet
                </div>
              ) : (
                // Always show only the first 3 conversations for the current workbook in the main view
                getWorkbookSessions(sessions).slice(0, 3).map(session => (
                  <div 
                    key={session.id} 
                    className={`flex justify-between items-center py-2 px-3 ${currentSession === session.id ? 'bg-gray-800/40' : 'hover:bg-black/20'} rounded-lg cursor-pointer border-l-2 border-black/40`}
                    onClick={() => loadSession(session.id)}
                  >
                    <div className="flex-grow text-white/80" style={{ 
                      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                      fontSize: '12px' 
                    }}>{session.title}</div>
                    <div className="ml-2" style={{ 
                      fontSize: '10px', 
                      opacity: 0.4, 
                      color: 'white', 
                      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                    }}>{formatTimeAgo(session.lastUpdated)}</div>
                  </div>
                ))
              )}
              
              {/* Show more button - always visible when there are more than 3 conversations for the current workbook */}
              {getWorkbookSessions(sessions).length > 3 && (
                <button
                  onClick={() => setShowPastConversationsView(true)}
                  className="text-gray-400 hover:text-gray-300 text-xs py-2 px-3 rounded-md hover:bg-black/20 w-full text-center transition-colors"
                  style={{ 
                    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
                  }}
                >
                  Show more ({getWorkbookSessions(sessions).length - 3} more)
                </button>
              )}
            </div>
          </div>
        </div>
      ) : (
        // Chat interface with Tailwind CSS styling
        <>
          {/* Messages area */}
          <div 
            ref={chatContainerRef} 
            className="flex-1 overflow-y-auto p-4 space-y-4"
          >
            {messages.map((message, index) => {
              // For status messages, render a custom traffic light indicator
              if (message.role === 'status') {
                // Determine traffic light color based on status
                const status = message.status || StatusType.Idle;
                let bgColor = '';
                let animate = false;
                
                // Add console logging to debug
                console.log('Status message:', message.content, 'with status:', status);
                
                // Use direct color values instead of Tailwind classes
                switch(status) {
                  case StatusType.Pending:
                  case StatusType.Idle:
                    bgColor = '#eab308'; // yellow-500
                    animate = true;
                    break;
                  case StatusType.Success:
                    bgColor = '#22c55e'; // green-500
                    break;
                  case StatusType.Error:
                    bgColor = '#ef4444'; // red-500
                    break;
                  default:
                    bgColor = '#6b7280'; // gray-500
                }
                
                return (
                  <div key={index} className="mb-2">
                    <div className="flex items-center gap-3 rounded-md px-3 py-2 shadow-md bg-gray-800/70 border border-gray-700">
                      {/* Pure traffic light circle indicator with direct styling */}
                      <div className="relative flex items-center justify-center">
                        <div 
                          className={`rounded-full ${animate ? 'animate-pulse' : ''}`} 
                          style={{ 
                            backgroundColor: bgColor,
                            width: '16px', 
                            height: '16px', 
                            boxShadow: '0 0 10px rgba(255,255,255,0.3)' 
                          }}
                        />
                      </div>
                      
                      {/* Status message */}
                      <span className="text-white" style={{ 
                        fontSize: 'clamp(10px, 1.25vw, 12px)', 
                        fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                      }}>
                        {message.content}
                      </span>
                    </div>
                  </div>
                );
              }
              
              // For user messages
              if (message.role === 'user') {
                return (
                  <div key={index} className="bg-gray-800/50 rounded-lg p-3 shadow-sm">
                    <div className="flex items-center">
                      <div className="bg-blue-500 text-white font-medium px-2 py-0.5 rounded mr-2 flex-shrink-0" style={{ 
                        fontSize: 'clamp(9px, 1vw, 11px)', 
                        fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                        border: '1px solid #3b82f6',
                        boxShadow: '0 0 0 1px rgba(59, 130, 246, 0.5)'
                      }}>
                        Me
                      </div>
                      <div 
                        className="text-white/90 rounded px-2 py-1" 
                        style={{ 
                          fontSize: 'clamp(11px, 1.5vw, 13px)', 
                          fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                          backgroundColor: 'rgba(75, 85, 99, 0.3)', // Light grey background (tailwind gray-600 with opacity)
                          width: '100%'
                        }}
                      >
                        {message.content}
                      </div>
                    </div>
                  </div>
                );
              }
              
              // For assistant messages
              if (message.role === 'assistant') {
                return (
                  <div key={index} className="rounded-lg p-3">
                    <div className="text-white/90 whitespace-pre-wrap" style={{ 
                      fontSize: 'clamp(11px, 1.5vw, 13px)', 
                      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                    }}>
                      {message.isStreaming ? (
                        <TypewriterEffect text={message.content || 'Thinking...'} speed={20} loop={false} />
                      ) : (
                        <div className="markdown-content">
                          <ReactMarkdown remarkPlugins={[remarkGfm]}>
                            {message.content}
                          </ReactMarkdown>
                        </div>
                      )}
                    </div>
                  </div>
                );
              }
              
              // For system messages (errors, etc.)
              return (
                <div key={index} className="bg-gray-900/60 rounded-lg p-3 text-red-400 shadow-sm" style={{ 
                  fontSize: 'clamp(11px, 1.5vw, 13px)', 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                }}>
                  {message.content}
                </div>
              );
            })}
            
            {/* Loading indicator */}
            {isLoading && !messages.some(m => m.isStreaming) && (
              <div className="flex items-center justify-center py-4">
                <div className="animate-pulse text-white/70">Processing...</div>
              </div>
            )}
          </div>
          
          {/* Extremely compact approval mode indicator */}
          <div className="flex items-center justify-end px-2 py-0.5 border-b border-gray-800">
            <div 
              className="flex items-center cursor-pointer"
              onClick={() => {
                const newState = !approvalEnabled;
                if (commandInterpreter) {
                  commandInterpreter.setRequireApproval(newState);
                }
                setApprovalEnabled(newState);
              }}
            >
              <div 
                className="h-2 w-2 rounded-full mr-1"
                style={{ backgroundColor: approvalEnabled ? '#3b82f6' : '#6b7280' }}
              />
              <span 
                style={{ 
                  fontSize: '8px', 
                  color: approvalEnabled ? '#d1d5db' : '#9ca3af',
                  fontFamily: 'monospace',
                  letterSpacing: '-0.5px'
                }}
              >
                {approvalEnabled ? 'Approval Mode' : 'Auto Mode'}
              </span>
            </div>
          </div>
          
          {/* Pending Changes Bar */}
          {approvalEnabled && pendingChanges.length > 0 && (
            <PendingChangesBar
              pendingChanges={pendingChanges}
              onAcceptAll={handleAcceptAll}
              onRejectAll={handleRejectAll}
              onAcceptChange={handleAcceptChange}
              onRejectChange={handleRejectChange}
            />
          )}
          
          {/* Input area */}
          <div className="bg-transparent p-2 mx-3 mb-3 rounded-lg">
            <div className="relative flex items-center border border-gray-700 hover:border-gray-500 focus-within:border-blue-500 rounded-md overflow-hidden">
              <textarea
                ref={textareaRef}
                className="flex-grow bg-transparent px-3 py-2 pr-10 text-white placeholder-white/50 resize-none outline-none min-h-[36px]"
                placeholder="Ask anything"
                value={userInput}
                onChange={handleInputChange}
                onKeyDown={handleKeyPress}
                disabled={isLoading || !servicesReady}
                rows={1}
                style={{ 
                  fontSize: 'clamp(11px, 1.5vw, 13px)',
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
                }}
              />
              <button
                className={`absolute right-3 p-1 rounded-md transition-all duration-300 flex items-center justify-center ${isLoading ? 'bg-red-600 animate-pulse' : 'text-gray-500 disabled:text-gray-700 hover:text-white hover:bg-gray-700/30'}`}
                onClick={handleSendMessage}
                disabled={!userInput.trim() || isLoading || !servicesReady}
                aria-label={isLoading ? "Processing" : "Send message"}
                type="button"
                style={{ 
                  width: '24px', 
                  height: '24px',
                  transition: 'all 0.3s ease'
                }}
              >
                {isLoading ? (
                  // Pulsing red square when loading
                  <div className="w-3 h-3 bg-white rounded-sm"></div>
                ) : (
                  // Arrow icon when not loading
                  <svg style={{ width: '16px', height: '16px' }} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M5 12H19M19 12L13 6M19 12L13 18" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                )}
              </button>
            </div>
          </div>
          {/* Removed the New Conversation button */}
        </>
      )}
    </div>
  );
};

export default TailwindFinancialModelChat;
