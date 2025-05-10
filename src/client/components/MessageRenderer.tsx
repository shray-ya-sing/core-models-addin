// src/client/components/chat/MessageRenderer.tsx
import React from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { ChatMessage } from '../models/ConversationModels';
import { TypewriterEffect } from './TypewriterEffect';
import { StatusType } from './StatusIndicator';

interface MessageRendererProps {
  message: ChatMessage;
}

const MessageRenderer: React.FC<MessageRendererProps> = ({ message }) => {
  // For status messages
  if (message.role === 'status') {
    const status = message.status || StatusType.Idle;
    let bgColor = '';
    let animate = false;
    
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
      <div className="mb-2">
        <div className="flex items-center gap-3 rounded-md px-3 py-2 shadow-md bg-gray-800/70 border border-gray-700">
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
          
          <span className="text-white" style={{ 
            fontSize: 'clamp(10px, 1.25vw, 12px)', 
            fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif' 
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
      <div className="bg-gray-800/50 rounded-lg p-3 shadow-sm">
        <div className="flex items-center">
          <div className="bg-blue-500 text-white font-medium px-2 py-0.5 rounded mr-2 flex-shrink-0" style={{ 
            fontSize: 'clamp(9px, 1vw, 11px)', 
            fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
            border: '1px solid #3b82f6',
            boxShadow: '0 0 0 1px rgba(59, 130, 246, 0.5)'
          }}>
            Me
          </div>
          <div 
            className="text-white/90 rounded px-2 py-1" 
            style={{ 
              fontSize: 'clamp(11px, 1.5vw, 13px)', 
              fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
              backgroundColor: 'rgba(75, 85, 99, 0.3)',
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
      <div className="rounded-lg p-3">
        <div className="text-white/90 whitespace-pre-wrap" style={{ 
          fontSize: 'clamp(11px, 1.5vw, 13px)', 
          fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif' 
        }}>
          {message.isStreaming ? (
            <div>
              <span>{message.content}</span>
              <span className="animate-pulse">▋</span>
            </div>
          ) : (
            <div className="markdown-content fix-numbered-lists">
              <ReactMarkdown
                remarkPlugins={[remarkGfm]}
              >
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
    <div className="bg-gray-900/60 rounded-lg p-3 text-red-400 shadow-sm" style={{ 
      fontSize: 'clamp(11px, 1.5vw, 13px)', 
      fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif' 
    }}>
      {message.content}
    </div>
  );
};

export default MessageRenderer;