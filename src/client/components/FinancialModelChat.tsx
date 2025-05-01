import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { v4 as uuidv4 } from 'uuid';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { ClientCommandManager } from '../services/ClientCommandManager';
import { ClientCommandExecutor } from '../services/ClientCommandExecutor';
import { ClientWorkbookStateManager } from '../services/ClientWorkbookStateManager';
import { ClientSpreadsheetCompressor } from '../services/ClientSpreadsheetCompressor';
import { ClientAnthropicService } from '../services/ClientAnthropicService';
import { ClientKnowledgeBaseService, KnowledgeBaseStatus, KnowledgeBaseEvent } from '../services/ClientKnowledgeBaseService';
import { ClientQueryProcessor, QueryProcessorResult } from '../services/ClientQueryProcessor';
import { Command, CommandStatus } from '../models/CommandModels';
import { ProcessStatusManager, ProcessStatus, ProcessStage } from '../models/ProcessStatusModels';
import { TypewriterEffect } from './TypewriterEffect';
import StatusIndicator, { StatusType } from './StatusIndicator';
import ProcessStatusTracker from './ProcessStatusTracker';
import { getStageName } from './ProcessStatusTracker';
import config from '../config';

// Create a component for the spinner animation
const Spinner = () => {
  // Add the keyframes to the document when component mounts
  useEffect(() => {
    // Only run in browser environment
    if (typeof document !== 'undefined') {
      const styleElement = document.createElement('style');
      styleElement.textContent = `
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
      `;
      document.head.appendChild(styleElement);
      
      return () => {
        // Clean up the style element when component unmounts
        document.head.removeChild(styleElement);
      };
    }
    return undefined; // Return undefined for non-browser environments
  }, []);
  
  return (
    <div style={{
      width: '12px',
      height: '12px',
      borderRadius: '50%',
      border: '2px solid #0078d4',
      borderTopColor: 'transparent',
      WebkitAnimation: 'spin 1s linear infinite',
      animation: 'spin 1s linear infinite'
    }}></div>
  );
};

// Message type definition
interface ChatMessage {
  role: 'user' | 'assistant' | 'system' | 'status';
  content: string;
  isStreaming?: boolean; // Flag to indicate if this message is currently being streamed
  status?: StatusType; // Status type for status messages
  stage?: string; // Stage identifier for status messages (e.g., ProcessStage or 'kb')
}

// Styles for the component
const styles = {
  container: {
    display: 'flex',
    flexDirection: 'column' as const,
    height: '100%',
    backgroundColor: '#000000',
    color: '#ffffff'
  },
  generatingIndicator: {
    padding: '8px 16px',
    color: '#999999',
    textAlign: 'center' as const,
    fontSize: '14px',
    fontStyle: 'italic' as const
  },
  welcomeScreen: {
    display: 'flex',
    flexDirection: 'column' as const,
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
    padding: '16px'
  },
  welcomeMessage: {
    fontSize: '24px',
    fontWeight: 'bold',
    marginBottom: '32px',
    textAlign: 'center' as const
  },
  centeredInputContainer: {
    display: 'flex',
    width: '100%',
    maxWidth: '600px'
  },
  chatContainer: {
    flex: 1,
    overflowY: 'auto' as const,
    padding: '16px',
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '8px',
    scrollbarWidth: 'thin' as const,
    scrollbarColor: '#333333 #121212'
  },
  inputContainer: {
    display: 'flex',
    padding: '16px',
    borderTop: '1px solid #333333',
    backgroundColor: '#121212'
  },
  input: {
    flex: 1,
    padding: '8px 12px',
    borderRadius: '4px',
    border: '1px solid #333333',
    backgroundColor: '#1e1e1e',
    color: '#ffffff',
    outline: 'none',
    resize: 'none' as const,
    minHeight: '20px',
    maxHeight: '120px',
    fontFamily: 'inherit'
  },
  sendButton: {
    marginLeft: '8px',
    padding: '8px 12px',
    backgroundColor: '#0078d4',
    color: 'white',
    border: 'none',
    borderRadius: '4px',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },
  message: {
    padding: '12px',
    borderRadius: '4px',
    maxWidth: '80%',
    wordBreak: 'break-word' as const
  },
  userMessage: {
    backgroundColor: '#333333',
    color: '#ffffff',
    alignSelf: 'flex-start',
    borderTopRightRadius: '0',
    display: 'flex',
    alignItems: 'flex-start'
  },
  assistantMessage: {
    backgroundColor: 'transparent',
    color: '#ffffff',
    alignSelf: 'flex-start',
    borderTopLeftRadius: '0'
  },
  systemMessage: {
    backgroundColor: '#0078d4',
    color: '#ffffff',
    alignSelf: 'center',
    fontSize: '0.9em',
    padding: '8px 12px'
  },
  userIcon: {
    backgroundColor: '#555555',
    color: '#ffffff',
    borderRadius: '50%',
    width: '24px',
    height: '24px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '0.8em',
    marginRight: '8px',
    flexShrink: 0
  },
  commandPanel: {
    marginTop: '8px',
    padding: '8px',
    backgroundColor: '#1e1e1e',
    borderRadius: '4px',
    fontSize: '0.9em'
  },
  statusIndicator: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '4px 8px',
    borderRadius: '4px',
    fontSize: '0.8em',
    marginBottom: '8px'
  },
  statusConnected: {
    backgroundColor: '#107C10',
    color: 'white'
  },
  statusDisconnected: {
    backgroundColor: '#D83B01',
    color: 'white'
  },
  statusConnecting: {
    backgroundColor: '#FFB900',
    color: 'black'
  },
  reconnectButton: {
    backgroundColor: '#0078d4',
    color: 'white',
    border: 'none',
    borderRadius: '4px',
    padding: '4px 8px',
    fontSize: '0.8em',
    cursor: 'pointer',
    marginLeft: '8px'
  },
  // Markdown styles
  markdownContent: {
    width: '100%',
    lineHeight: '1.5',
    fontSize: '14px'
  }
};

// Add global styles for markdown elements
if (typeof document !== 'undefined') {
  // Only run in browser environment
  const styleElement = document.createElement('style');
  styleElement.textContent = `
    .markdown-content h1 {
      font-size: 1.8em;
      margin-top: 1em;
      margin-bottom: 0.5em;
      font-weight: 600;
      border-bottom: 1px solid #333;
      padding-bottom: 0.2em;
    }
    
    .markdown-content h2 {
      font-size: 1.5em;
      margin-top: 0.8em;
      margin-bottom: 0.4em;
      font-weight: 600;
    }
    
    .markdown-content h3 {
      font-size: 1.3em;
      margin-top: 0.6em;
      margin-bottom: 0.3em;
      font-weight: 600;
    }
    
    .markdown-content h4, .markdown-content h5, .markdown-content h6 {
      font-size: 1.1em;
      margin-top: 0.5em;
      margin-bottom: 0.2em;
      font-weight: 600;
    }
    
    .markdown-content p {
      margin-bottom: 1em;
    }
    
    .markdown-content ul, .markdown-content ol {
      margin-bottom: 1em;
      padding-left: 2em;
    }
    
    .markdown-content li {
      margin-bottom: 0.3em;
    }
    
    .markdown-content code {
      background-color: #2d2d2d;
      padding: 0.2em 0.4em;
      border-radius: 3px;
      font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
      font-size: 0.9em;
    }
    
    .markdown-content pre {
      background-color: #1e1e1e;
      border-radius: 4px;
      padding: 1em;
      margin: 1em 0;
      overflow-x: auto;
    }
    
    .markdown-content pre code {
      background-color: transparent;
      padding: 0;
      border-radius: 0;
      font-size: 0.9em;
      color: #e6e6e6;
    }
    
    .markdown-content blockquote {
      border-left: 4px solid #555;
      padding-left: 1em;
      margin-left: 0;
      margin-right: 0;
      font-style: italic;
      color: #aaa;
    }
    
    .markdown-content table {
      border-collapse: collapse;
      width: 100%;
      margin-bottom: 1em;
      color: #eee;
    }
    
    .markdown-content th, .markdown-content td {
      border: 1px solid #444;
      padding: 0.5em;
      text-align: left;
    }
    
    .markdown-content th {
      background-color: #333;
    }
    
    .markdown-content tr:nth-child(even) {
      background-color: #292929;
    }
    
    .markdown-content a {
      color: #4da3ff;
      text-decoration: none;
    }
    
    .markdown-content a:hover {
      text-decoration: underline;
    }
    
    .markdown-content img {
      max-width: 100%;
      height: auto;
    }
    
    .markdown-content hr {
      border: none;
      border-top: 1px solid #444;
      margin: 1.5em 0;
    }
  `;
  document.head.appendChild(styleElement);
}

export const FinancialModelChat: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [userInput, setUserInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [hasUserSentMessage, setHasUserSentMessage] = useState(false);
  const [commandManager, setCommandManager] = useState<ClientCommandManager | null>(null);
  const [anthropicService, setAnthropicService] = useState<ClientAnthropicService | null>(null);
  const [knowledgeBaseService, setKnowledgeBaseService] = useState<ClientKnowledgeBaseService | null>(null);
  const [workbookStateManager, setWorkbookStateManager] = useState<ClientWorkbookStateManager | null>(null);
  const [spreadsheetCompressor, setSpreadsheetCompressor] = useState<ClientSpreadsheetCompressor | null>(null);
  const [queryProcessor, setQueryProcessor] = useState<ClientQueryProcessor | null>(null);
  const [servicesReady, setServicesReady] = useState(false);
  const [kbStatus, setKbStatus] = useState<{ status: StatusType; message: string } | null>(null);
  const [streamingResponse, setStreamingResponse] = useState<string>('');
  const [isStreaming, setIsStreaming] = useState<boolean>(false);
  const chatContainerRef = useRef<HTMLDivElement>(null);
  
  // Initialize client-side services
  useEffect(() => {
    const initializeClientServices = async () => {
      try {
        // Log initialization
        console.log('%c Initializing client services...', 'background: #222; color: #bada55; font-size: 14px');
        
        // Check for API key
        if (!config.anthropicApiKey) {
          console.warn('No Anthropic API key found in configuration');
        }
        
        // Initialize services
        const stateManager = new ClientWorkbookStateManager();
        const compressor = new ClientSpreadsheetCompressor();
        const anthropic = new ClientAnthropicService(config.anthropicApiKey);
        const knowledgeBase = new ClientKnowledgeBaseService(config.knowledgeBaseApiUrl);
        
        // Initialize command execution system
        const commandExecutor = new ClientCommandExecutor(stateManager);
        // Pass the WorkbookStateManager to the CommandManager for cache invalidation
        const manager = new ClientCommandManager(commandExecutor, stateManager);
        
        // Register for command updates
        const unsubscribeCommandUpdate = manager.onCommandUpdate((command) => {
          console.log('Command update received:', command);
          
          // If the command is completed, we can add a message to the chat
          if (command.status === CommandStatus.Completed) {
            setMessages(prev => [...prev, { 
              role: 'system', 
              content: `Command "${command.description}" completed successfully.` 
            }]);
          } else if (command.status === CommandStatus.Failed) {
            setMessages(prev => [...prev, { 
              role: 'system', 
              content: `Command "${command.description}" failed: ${command.error || 'Unknown error'}` 
            }]);
          }
        });
        
        // Create query processor
        const processor = new ClientQueryProcessor({
          anthropic,
          kbService: knowledgeBase,
          workbookStateManager: stateManager,
          compressor,
          commandManager: manager,
          useAdvancedChunkLocation: true
        });
        
        // Set up knowledge base event listener
        const unsubscribeKbEvents = knowledgeBase.addEventListener((event: KnowledgeBaseEvent) => {
          console.log('Knowledge Base Event:', event);
          
          // Map KB status to StatusIndicator status
          let statusType: StatusType;
          switch (event.status) {
            case KnowledgeBaseStatus.Searching:
              statusType = StatusType.Pending;
              break;
            case KnowledgeBaseStatus.Success:
              statusType = StatusType.Success;
              break;
            case KnowledgeBaseStatus.Error:
              statusType = StatusType.Error;
              break;
            default:
              statusType = StatusType.Idle;
          }
          
          // Add KB status as a message in the chat
          setMessages(prev => {
            // Find the last user message index
            const lastUserMsgIndex = [...prev].reverse().findIndex(m => m.role === 'user');
            
            // If we found a user message
            if (lastUserMsgIndex >= 0) {
              const actualIndex = prev.length - 1 - lastUserMsgIndex;
              
              // First, remove any pending messages for Knowledge Base if we're getting a success/error
              let newMessages = [...prev];
              if (statusType === StatusType.Success || statusType === StatusType.Error) {
                newMessages = newMessages.filter(m => !(m.role === 'status' && m.stage === 'kb' && m.status === StatusType.Pending));
              }
              
              // Now find if we have any existing status message for Knowledge Base after filtering
              const existingStatusIndex = newMessages.findIndex(m => m.role === 'status' && m.stage === 'kb' && m.status === statusType);
              
              // Create the status message
              const statusMessage = {
                role: 'status' as const,
                content: event.message,
                status: statusType,
                stage: 'kb'
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
          
          // Keep the KB status state for backward compatibility
          setKbStatus({ status: statusType, message: event.message });
          
          // Keep KB status state for backward compatibility
          if (event.status === KnowledgeBaseStatus.Success) {
            setTimeout(() => {
              setKbStatus(null);
            }, 5000);
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
          
          // We're keeping all status messages visible now
          // No auto-removal for successful status messages
        });
        
        // Store services in state
        setWorkbookStateManager(stateManager);
        setSpreadsheetCompressor(compressor);
        setAnthropicService(anthropic);
        setKnowledgeBaseService(knowledgeBase);
        setCommandManager(manager);
        setQueryProcessor(processor);
        console.log('%c All services initialized successfully!', 'background: #222; color: #2ecc71; font-size: 14px');
        setServicesReady(true);
        
        // No welcome message in the chat history anymore
        
        // Return cleanup function
        return () => {
          unsubscribeCommandUpdate();
          unsubscribeKbEvents();
          unsubscribeProcessEvents();
        };
      } catch (error) {
        console.error('Error initializing client services:', error);
        setMessages(prev => [...prev, { 
          role: 'system', 
          content: `Error initializing services: ${error.message}` 
        }]);
        return () => {};
      }
    };
    
    const cleanup = initializeClientServices();
    
    // Cleanup function
    return () => {
      if (cleanup) {
        cleanup.then(cleanupFn => cleanupFn());
      }
    };
  }, []);

  // Auto-scroll to bottom when messages change
  useEffect(() => {
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);

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
      
      // If there's a command in the response, show a message about it
      if (result.command) {
        setMessages(prev => [...prev, { 
          role: 'system', 
          content: `Executing command: ${result.command.description}` 
        }]);
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

  // Render the component
  return (
    <div style={styles.container}>
      {!hasUserSentMessage ? (
        // Initial welcome screen
        <div style={styles.welcomeScreen}>
          <div style={styles.welcomeMessage}>Build faster than ever with Cori</div>
          <div style={styles.centeredInputContainer}>
            <textarea
              style={styles.input}
              placeholder="Type your message..."
              value={userInput}
              onChange={handleInputChange}
              onKeyPress={handleKeyPress}
              disabled={!servicesReady}
              rows={1}
            />
            <button 
              style={styles.sendButton} 
              onClick={handleSendMessage}
              disabled={!userInput.trim() || isLoading || !servicesReady}
            >
              {isLoading ? <Spinner /> : '→'}
            </button>
          </div>
        </div>
      ) : (
        // Chat interface after first message
        <>
          <div ref={chatContainerRef} style={styles.chatContainer}>
            {messages.map((message, index) => {
              // For status messages, render a StatusIndicator
              if (message.role === 'status') {
                return (
                  <div key={index} style={{ marginBottom: '8px' }}>
                    <StatusIndicator 
                      status={message.status || StatusType.Idle}
                      message={message.content}
                      autoHide={false}
                    />
                  </div>
                );
              }
              
              // For other message types, render as before
              return (
                <div 
                  key={index} 
                  style={{
                    ...styles.message,
                    ...(message.role === 'user' ? styles.userMessage : 
                       message.role === 'assistant' ? styles.assistantMessage : 
                       styles.systemMessage)
                  }}
                >
                  {message.role === 'user' && (
                    <div style={styles.userIcon}>Me</div>
                  )}
                  {message.role === 'assistant' ? (
                    <div className="markdown-content">
                      <ReactMarkdown remarkPlugins={[remarkGfm]}>
                        {message.content}
                      </ReactMarkdown>
                    </div>
                  ) : (
                    message.content
                  )}
                </div>
              );
            })}
            
            {/* Legacy loading indicator - we'll eventually phase this out */}
            {isLoading && (
              <div style={styles.generatingIndicator}>
                <TypewriterEffect text="Processing" speed={120} loop={true} />
              </div>
            )}
          </div>
          
          <div style={styles.inputContainer}>
            <textarea
              style={styles.input}
              placeholder="Type your message..."
              value={userInput}
              onChange={handleInputChange}
              onKeyPress={handleKeyPress}
              disabled={!servicesReady}
              rows={1}
            />
            <button 
              style={styles.sendButton} 
              onClick={handleSendMessage}
              disabled={!userInput.trim() || isLoading || !servicesReady}
            >
              {isLoading ? <Spinner /> : '→'}
            </button>
          </div>
        </>
      )}
    </div>
  );
};

export default FinancialModelChat;
