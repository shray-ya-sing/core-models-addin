import * as React from 'react';
import { RefObject, useState } from 'react';
import { ConversationSession } from '../models/ConversationModels';

/**
 * Props for the StartNewConversation component
 */
interface StartNewConversationProps {
  textareaRef?: RefObject<HTMLTextAreaElement>;
  onMessageSent?: (message: string) => void;
  servicesReady?: boolean;
  sessions?: ConversationSession[];
  currentSession?: string;
  getWorkbookSessions?: (allSessions: ConversationSession[]) => ConversationSession[];
  loadSession?: (sessionId: string) => void;
  formatTimeAgo?: (timestamp: number) => string;
  setShowPastConversationsView?: (show: boolean) => void;
}
/**
 * Component for the new conversation view
 */
const StartNewConversation: React.FC<StartNewConversationProps> = ({
  textareaRef: externalTextareaRef,
  onMessageSent,
  servicesReady = true,
  sessions = [],
  currentSession = '',
  getWorkbookSessions = (sessions) => sessions,
  loadSession = () => {},
  formatTimeAgo = (timestamp) => {
    const now = Date.now();
    const secondsAgo = Math.floor((now - timestamp) / 1000);
    
    if (secondsAgo < 60) return `${secondsAgo}s`;
    if (secondsAgo < 3600) return `${Math.floor(secondsAgo / 60)}m`;
    if (secondsAgo < 86400) return `${Math.floor(secondsAgo / 3600)}h`;
    return `${Math.floor(secondsAgo / 86400)}d`;
  },
  setShowPastConversationsView = () => {}
}) => {
  // Internal state for the component
  const [userInput, setUserInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const internalTextareaRef = React.useRef<HTMLTextAreaElement>(null);
  
  // Use external ref if provided, otherwise use internal ref
  const textareaRef = externalTextareaRef || internalTextareaRef;
  
  // Handle input change
  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setUserInput(e.target.value);
  };
  
  // Handle key press (Enter to send)
  const handleKeyPress = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleStartNewConversation();
    }
  };
  
  // Handle starting a new conversation
  const handleStartNewConversation = () => {
    if (!userInput.trim() || isLoading || !servicesReady) return;
    
    setIsLoading(true);
    
    // Call the onMessageSent callback with the user's message
    if (onMessageSent) {
      onMessageSent(userInput.trim());
    }
    
    // Clear the input field
    setUserInput('');
    setIsLoading(false);
  };
  return (
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
                fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif'
              }} 
              alt="Cori Logo" 
            />  
          </div>
          
          {/* Description with fixed font size */}
          <p 
            className="text-white/70 leading-relaxed text-left mb-6 max-w-sm" 
            style={{ 
              fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
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
              onClick={handleStartNewConversation}
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
            fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
            fontSize: '12px'
          }}>Past Conversations</h2>
        </div>
        <div className="flex flex-col space-y-0.5">
          {sessions.length === 0 ? (
            <div className="text-gray-500 text-center py-2" style={{ 
              fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
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
                onClick={() => {
                  try {
                    loadSession(session.id);
                  } catch (error) {
                    console.error('Error loading session from main view:', error);
                  }
                }}
              >
                <div className="flex-grow text-white/80" style={{ 
                  fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
                  fontSize: '12px' 
                }}>{session.title}</div>
                <div className="ml-2" style={{ 
                  fontSize: '10px', 
                  opacity: 0.4, 
                  color: 'white', 
                  fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif' 
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
                fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif'
              }}
            >
              Show more ({getWorkbookSessions(sessions).length - 3} more)
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default StartNewConversation;
