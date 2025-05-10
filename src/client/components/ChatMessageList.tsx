// src/client/components/chat/ChatMessageList.tsx
import React, { useRef, useEffect } from 'react';
import { ChatMessage } from '../models/ConversationModels';
import MessageRenderer from './MessageRenderer';

interface ChatMessageListProps {
  messages: ChatMessage[];
  isLoading: boolean;
  isStreaming?: boolean;
}

const ChatMessageList: React.FC<ChatMessageListProps> = ({ messages, isLoading, isStreaming = false }) => {
  const chatContainerRef = useRef<HTMLDivElement>(null);
  
  // Auto-scroll to bottom when messages change
  useEffect(() => {
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);

  // Group messages into conversations (user + status + assistant)
  const groupMessages = (msgs: ChatMessage[]) => {
    const groups: ChatMessage[][] = [];
    let currentGroup: ChatMessage[] = [];
    
    // Process messages in chronological order (oldest first)
    [...msgs].forEach(message => {
      // Start a new group when we encounter a user message
      if (message.role === 'user' && currentGroup.length > 0) {
        groups.push([...currentGroup]);
        currentGroup = [];
      }
      
      // Add the current message to the group
      currentGroup.push(message);
    });
    
    // Add the last group if it has messages
    if (currentGroup.length > 0) {
      groups.push(currentGroup);
    }
    
    return groups; // Return groups in chronological order
  };

  const messageGroups = groupMessages(messages);

  return (
    <div 
      ref={chatContainerRef} 
      className="flex-1 overflow-y-auto p-4 space-y-4"
    >
      {messageGroups.map((group, groupIndex) => (
        <div key={groupIndex} className="space-y-4">
          {group.map((message, messageIndex) => (
            <MessageRenderer key={`${groupIndex}-${messageIndex}`} message={message} />
          ))}
        </div>
      ))}
      
      {/* Loading indicator */}
      {(isLoading || isStreaming) && !messages.some(m => m.isStreaming) && (
        <div className="flex items-center justify-center py-4">
          <div className="animate-pulse text-white/70">
            {isStreaming ? "Generating response..." : "Processing..."}
          </div>
        </div>
      )}
    </div>
  );
};

export default ChatMessageList;