import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { ChatMessage, Attachment } from '../models/ConversationModels';
import ChatMessageList from './ChatMessageList';
import ChatInputArea from './ChatInputArea';
import { useChatServices } from '../hooks/useChatServices';
import useConversationState from '../hooks/useConversationState';

interface TailwindFinancialModelChatProps {
  initialMessage?: string;
  resetChat?: boolean;
}

const TailwindFinancialModelChat: React.FC<TailwindFinancialModelChatProps> = ({ 
  initialMessage = '',
  resetChat = false
}) => {
  // Basic UI state
  const [isLoading, setIsLoading] = useState(false);
  const [userInput, setUserInput] = useState('');
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [isStreaming, setIsStreaming] = useState(false);
  const [streamingResponse, setStreamingResponse] = useState('');
  
  // Track processed messages to avoid duplicates
  const processedMessagesRef = useRef<Set<string>>(new Set());
  
  // Initialize services
  const {
    queryProcessor,
    servicesReady,
    currentWorkbookId
  } = useChatServices();
  
  // Initialize conversation state
  const {
    messages,
    setMessages,
    createNewSession,
    updateSessionMessages
  } = useConversationState({ currentWorkbookId });
  
  // Handle reset chat
  useEffect(() => {
    if (resetChat) {
      // Clear all messages
      setMessages([]);
      processedMessagesRef.current.clear();
      
      // Create a new session
      createNewSession();
    }
  }, [resetChat, setMessages, createNewSession]);
  
  // Process initial message
  useEffect(() => {
    const processInitialMessage = async () => {
      // Only process if we have an initialMessage that hasn't been processed yet
      if (initialMessage && 
          !processedMessagesRef.current.has(initialMessage) && 
          servicesReady && 
          !isLoading) {
        
        // Mark as processed to prevent duplicates
        processedMessagesRef.current.add(initialMessage);
        
        // Create a new user message
        const userMessage: ChatMessage = {
          role: 'user',
          content: initialMessage
        };
        
        // Set the message as the only message in the chat
        setMessages([userMessage]);
        
        // Set loading state
        setIsLoading(true);
        
        try {
          if (queryProcessor) {
            // Set up streaming state
            setIsStreaming(true);
            setStreamingResponse('');
            
            // Create a streaming response handler
            const handleStreamingResponse = (chunk: string) => {
              setStreamingResponse(prev => {
                const newResponse = prev + chunk;
                
                // Update messages with streaming response
                setMessages([
                  userMessage,
                  {
                    role: 'assistant',
                    content: newResponse,
                    isStreaming: true
                  }
                ]);
                
                return newResponse;
              });
            };
            
            // Process the query with streaming handler
            const result = await queryProcessor.processQuery(
              initialMessage,
              handleStreamingResponse, // Streaming handler
              [], // Empty chat history
              [] // No attachments
            );
            
            // Streaming complete, update with final message
            setIsStreaming(false);
            
            // Add the assistant response (final version)
            setMessages([
              userMessage,
              {
                role: 'assistant',
                content: result.assistantMessage,
                isStreaming: false
              }
            ]);
          }
        } catch (error) {
          console.error('Error processing message:', error);
          
          // Add error message
          setMessages([
            userMessage,
            {
              role: 'system',
              content: 'I encountered an error processing your request.'
            }
          ]);
        } finally {
          setIsLoading(false);
        }
      }
    };
    
    processInitialMessage();
  }, [initialMessage, servicesReady, isLoading, queryProcessor, setMessages]);
  
  // Handle sending a message
  const handleSendMessage = async () => {
    if (!userInput.trim() || !servicesReady || isLoading) return;
    
    // Create user message with any attachments
    const userMessage: ChatMessage = {
      role: 'user',
      content: userInput,
      attachments: attachments.length > 0 ? [...attachments] : undefined
    };
    
    // Add to messages
    setMessages(prev => [...prev, userMessage]);
    
    // Clear input and attachments
    setUserInput('');
    setAttachments([]);
    
    // Set loading state
    setIsLoading(true);
    
    try {
      if (queryProcessor) {
        // Set up streaming state
        setIsStreaming(true);
        setStreamingResponse('');
        
        // Create a streaming response handler
        const handleStreamingResponse = (chunk: string) => {
          setStreamingResponse(prev => {
            const newResponse = prev + chunk;
            
            // Check if we already have a streaming message in the current conversation
            setMessages(prev => {
              const newMessages = [...prev];
              
              // Find the last user message first to identify the current conversation
              const lastUserMsgIndex = [...newMessages].reverse().findIndex(m => m.role === 'user');
              
              if (lastUserMsgIndex === -1) {
                // No user message found, just append the assistant message
                newMessages.push({
                  role: 'assistant',
                  content: newResponse,
                  isStreaming: true
                });
                return newMessages;
              }
              
              // Convert to actual index (from the end)
              const actualUserIndex = newMessages.length - 1 - lastUserMsgIndex;
              
              // Look for a streaming assistant message after the last user message
              let assistantMsgIndex = -1;
              for (let i = actualUserIndex + 1; i < newMessages.length; i++) {
                if (newMessages[i].role === 'assistant' && newMessages[i].isStreaming === true) {
                  assistantMsgIndex = i;
                  break;
                }
              }
              
              if (assistantMsgIndex >= 0) {
                // Update existing streaming message in the current conversation
                newMessages[assistantMsgIndex] = {
                  ...newMessages[assistantMsgIndex],
                  content: newResponse
                };
              } else {
                // Add new streaming message after the last user message and any status messages
                let insertIndex = actualUserIndex + 1;
                
                // Skip past any status messages
                while (insertIndex < newMessages.length && newMessages[insertIndex].role === 'status') {
                  insertIndex++;
                }
                
                // Insert the new assistant message at the right position
                newMessages.splice(insertIndex, 0, {
                  role: 'assistant',
                  content: newResponse,
                  isStreaming: true
                });
              }
              
              return newMessages;
            });
            
            return newResponse;
          });
        };
        
        // Process the query with streaming handler
        const result = await queryProcessor.processQuery(
          userInput,
          handleStreamingResponse, // Streaming handler
          messages, // Current chat history
          attachments // Include any attachments
        );
        
        // Streaming complete
        setIsStreaming(false);
        
        // Update the final message without streaming flag
        setMessages(prev => {
          const newMessages = [...prev];
          const assistantMsgIndex = newMessages.findIndex(m => 
            m.role === 'assistant' && m.isStreaming === true
          );
          
          if (assistantMsgIndex >= 0) {
            newMessages[assistantMsgIndex] = {
              role: 'assistant',
              content: result.assistantMessage,
              isStreaming: false
            };
          } else {
            // Fallback if no streaming message was found
            newMessages.push({
              role: 'assistant',
              content: result.assistantMessage
            });
          }
          
          return newMessages;
        });
      }
    } catch (error) {
      console.error('Error processing message:', error);
      
      // Add error message
      setMessages(prev => [
        ...prev,
        {
          role: 'system',
          content: 'I encountered an error processing your request.'
        }
      ]);
    } finally {
      setIsLoading(false);
    }
  };
  
  // Handle file uploads
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (!event.target.files || event.target.files.length === 0) return;
    
    // Process each selected file
    Array.from(event.target.files).forEach(file => {
      // Check file size (limit to 5MB)
      if (file.size > 5 * 1024 * 1024) {
        console.warn(`File ${file.name} exceeds 5MB size limit and will be skipped`);
        return;
      }
      
      const reader = new FileReader();
      
      reader.onload = (e) => {
        if (!e.target || typeof e.target.result !== 'string') return;
        
        // Determine file type
        let fileType: 'image' | 'pdf' = 'image';
        if (file.type === 'application/pdf') {
          fileType = 'pdf';
        } else if (!file.type.startsWith('image/')) {
          console.warn(`Unsupported file type: ${file.type}`);
          return;
        }
        
        // Create attachment object
        const newAttachment: Attachment = {
          type: fileType,
          name: file.name,
          content: e.target.result,
          mimeType: file.type
        };
        
        // Add to attachments array
        setAttachments(prev => [...prev, newAttachment]);
      };
      
      reader.onerror = () => {
        console.error(`Error reading file: ${file.name}`);
      };
      
      // Read file as data URL
      reader.readAsDataURL(file);
    });
    
    // Reset the input field to allow selecting the same file again
    event.target.value = '';
  };
  
  // Remove attachment
  const removeAttachment = (index: number) => {
    setAttachments(prev => prev.filter((_, i) => i !== index));
  };
  
  return (
    <div className="flex flex-col h-full font-mono text-sm relative">
      {/* Messages area */}
      <ChatMessageList 
        messages={messages}
        isLoading={isLoading}
        isStreaming={isStreaming}
      />
      
      {/* Input area */}
      <ChatInputArea 
        userInput={userInput}
        setUserInput={setUserInput}
        handleSendMessage={handleSendMessage}
        isLoading={isLoading}
        servicesReady={servicesReady}
        attachments={attachments}
        handleFileChange={handleFileChange}
        removeAttachment={removeAttachment}
      />
    </div>
  );
};

export default TailwindFinancialModelChat;