// src/client/components/chat/ChatInputArea.tsx
import React, { useRef } from 'react';
import { Attachment } from '../models/ConversationModels';

interface ChatInputAreaProps {
  userInput: string;
  setUserInput: (input: string) => void;
  handleSendMessage: () => void;
  isLoading: boolean;
  servicesReady: boolean;
  attachments: Attachment[];
  handleFileChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  removeAttachment: (index: number) => void;
}

const ChatInputArea: React.FC<ChatInputAreaProps> = ({
  userInput,
  setUserInput,
  handleSendMessage,
  isLoading,
  servicesReady,
  attachments,
  handleFileChange,
  removeAttachment
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setUserInput(e.target.value);
  };

  const handleKeyPress = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  return (
    <div className="bg-transparent p-2 mx-3 mb-3 rounded-lg">
      {/* Attachment preview area */}
      {attachments.length > 0 && (
        <div className="flex flex-wrap gap-2 mb-2">
          {attachments.map((attachment, index) => (
            <div key={index} className="relative bg-gray-800 rounded-md p-1 flex items-center">
              <span className="text-xs text-white mr-1">
                {attachment.type === 'image' ? '🖼️' : '📄'} {attachment.name.length > 15 ? attachment.name.substring(0, 12) + '...' : attachment.name}
              </span>
              <button 
                onClick={() => removeAttachment(index)}
                className="text-gray-400 hover:text-red-400 ml-1"
                aria-label="Remove attachment"
              >
                <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"></path>
                </svg>
              </button>
            </div>
          ))}
        </div>
      )}
      
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
        
        {/* Attachment buttons */}
        <div className="flex items-center mr-1">
          {/* Image/PDF attachment buttons */}
          <button
            className="text-gray-500 hover:text-white p-1 rounded transition-colors"
            onClick={() => fileInputRef.current?.click()}
            disabled={isLoading}
            aria-label="Attach file"
            title="Attach file"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"></path>
            </svg>
          </button>
          
          {/* Hidden file input */}
          <input 
            type="file" 
            ref={fileInputRef} 
            className="hidden" 
            accept="image/*,.pdf"
            onChange={handleFileChange}
            disabled={isLoading}
          />
        </div>
        
        <button
          className={`p-1 rounded-md transition-all duration-300 flex items-center justify-center mr-2 ${isLoading ? 'bg-red-600 animate-pulse' : 'text-gray-500 disabled:text-gray-700 hover:text-white hover:bg-gray-700/30'}`}
          onClick={handleSendMessage}
          disabled={(!userInput.trim() && attachments.length === 0) || isLoading || !servicesReady}
          aria-label={isLoading ? "Processing" : "Send message"}
          type="button"
          style={{ 
            width: '24px', 
            height: '24px',
            transition: 'all 0.3s ease'
          }}
        >
          {isLoading ? (
            <div className="w-3 h-3 bg-white rounded-sm"></div>
          ) : (
            <svg style={{ width: '16px', height: '16px' }} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M5 12H19M19 12L13 6M19 12L13 18" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          )}
        </button>
      </div>
    </div>
  );
};

export default ChatInputArea;