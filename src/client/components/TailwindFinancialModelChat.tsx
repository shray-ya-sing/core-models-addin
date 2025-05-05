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
import { ClientKnowledgeBaseService } from '../services/ClientKnowledgeBaseService';
import { ClientQueryProcessor, QueryProcessorResult } from '../services/ClientQueryProcessor';
import { Command, CommandStatus } from '../models/CommandModels';
import { ProcessStatusManager, ProcessStatus, ProcessStage } from '../models/ProcessStatusModels';
import { TypewriterEffect } from './TypewriterEffect';
import StatusIndicator, { StatusType } from './StatusIndicator';
import ProcessStatusTracker from './ProcessStatusTracker';
import { getStageName } from './ProcessStatusTracker';
import config from '../config';
import { SendIcon, FileIcon, CheckIcon, AlertCircleIcon } from './icons';

// Message type definition
interface ChatMessage {
  role: 'user' | 'assistant' | 'system' | 'status';
  content: string;
  isStreaming?: boolean;
  status?: StatusType;
  stage?: string;
}

const TailwindFinancialModelChat: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [userInput, setUserInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [hasUserSentMessage, setHasUserSentMessage] = useState(false);
  const [servicesReady, setServicesReady] = useState(false);
  const [streamingResponse, setStreamingResponse] = useState<string>('');
  const [isStreaming, setIsStreaming] = useState<boolean>(false);
  
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

  // Status trackers
  const [statusManager] = useState<ProcessStatusManager>(() => ProcessStatusManager.getInstance());
  const [currentStatus, setCurrentStatus] = useState<ProcessStatus | null>(null);

  // Initialize services
  useEffect(() => {
    // Initialize all client services
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
                
        // Create query processor
        const processor = new ClientQueryProcessor({
          anthropic,
          kbService: knowledgeBase,
          workbookStateManager: stateManager,
          compressor,
          commandManager: manager,
          useAdvancedChunkLocation: true
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
        
        // If there's a command in the response, log it but don't show a message
        if (result.command) {
          console.log(`Executing command: ${result.command.description}`);
          // Removed system message
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

  return (
    <div className="flex flex-col h-full font-mono text-sm">
      {!hasUserSentMessage ? (
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
          <div className="w-full max-w-xl self-center">
            <h2 className="font-medium text-white mb-3" style={{ 
              fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
              fontSize: '12px'
            }}>Past Conversations</h2>
            <div className="flex flex-col space-y-0.5">
              {/* Sample conversations */}
              <div className="flex justify-between items-center py-2 px-3 hover:bg-black/20 rounded-lg cursor-pointer">
                <div className="flex-grow text-white/80" style={{ 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                  fontSize: '12px' 
                }}>Q1 Earnings Forecast Model</div>
                <div className="ml-2" style={{ 
                  fontSize: '10px', 
                  opacity: 0.4, 
                  color: 'white', 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                }}>38s</div>
              </div>
              <div className="flex justify-between items-center py-2 px-3 hover:bg-black/20 rounded-lg cursor-pointer">
                <div className="flex-grow text-white/80" style={{ 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                  fontSize: '12px' 
                }}>SaaS Company Valuation</div>
                <div className="ml-2" style={{ 
                  fontSize: '10px', 
                  opacity: 0.4, 
                  color: 'white', 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                }}>40m</div>
              </div>
              <div className="flex justify-between items-center py-2 px-3 hover:bg-black/20 rounded-lg cursor-pointer">
                <div className="flex-grow text-white/80" style={{ 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
                  fontSize: '12px' 
                }}>Merger Synergy Analysis</div>
                <div className="ml-2" style={{ 
                  fontSize: '10px', 
                  opacity: 0.4, 
                  color: 'white', 
                  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                }}>1h</div>
              </div>
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
                      <div className="text-white/90" style={{ 
                        fontSize: 'clamp(11px, 1.5vw, 13px)', 
                        fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
                      }}>
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
                className="absolute right-3 p-1 text-gray-500 disabled:text-gray-700 hover:text-white hover:bg-gray-700/30 rounded-md transition-colors flex items-center justify-center"
                onClick={handleSendMessage}
                disabled={!userInput.trim() || isLoading || !servicesReady}
                aria-label="Send message"
                type="button"
              >
                <svg style={{ width: '16px', height: '16px' }} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M5 12H19M19 12L13 6M19 12L13 18" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default TailwindFinancialModelChat;
