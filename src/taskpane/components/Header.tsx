import * as React from "react";
import { useState } from "react";
import { PlusIcon, ClockIcon, BookOpenIcon, MoreHorizontalIcon, XIcon, GitBranchIcon } from "lucide-react";

export interface HeaderProps {
  onNewConversation: () => void;
  onShowPastConversations: () => void;
  onShowVersionHistory?: () => void;
}

const Header: React.FC<HeaderProps> = ({ onNewConversation, onShowPastConversations, onShowVersionHistory }) => {
  const [showNewConversationTooltip, setShowNewConversationTooltip] = useState(false);
  const [showPastConversationsTooltip, setShowPastConversationsTooltip] = useState(false);
  const [showVersionHistoryTooltip, setShowVersionHistoryTooltip] = useState(false);
  
  return (
    <header className="flex justify-between items-center py-1 px-2 bg-transparent z-10">
      <div className="flex items-center text-gray-300 font-medium text-xs">
        Cori
      </div>
      <div className="flex items-center gap-2">
        {/* Version History (Git Branch) Icon */}
        <div className="relative">
          <button 
            onClick={onShowVersionHistory}
            onMouseEnter={() => setShowVersionHistoryTooltip(true)}
            onMouseLeave={() => setShowVersionHistoryTooltip(false)}
            className="flex items-center justify-center"
            aria-label="Version History"
          >
            <GitBranchIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-white hover:filter hover:brightness-200 transition-all duration-150" />
          </button>
          {showVersionHistoryTooltip && (
            <div className="absolute -bottom-8 left-1/2 transform -translate-x-1/2 px-2 py-1 bg-gray-800 text-gray-200 text-xs rounded whitespace-nowrap">
              Version History
            </div>
          )}
        </div>
        
        <div className="relative">
          <button 
            onClick={onNewConversation}
            onMouseEnter={() => setShowNewConversationTooltip(true)}
            onMouseLeave={() => setShowNewConversationTooltip(false)}
            className="flex items-center justify-center"
            aria-label="New Conversation"
          >
            <PlusIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-white hover:filter hover:brightness-200 transition-all duration-150" />
          </button>
          {showNewConversationTooltip && (
            <div className="absolute -bottom-8 left-1/2 transform -translate-x-1/2 px-2 py-1 bg-gray-800 text-gray-200 text-xs rounded whitespace-nowrap">
              New Conversation
            </div>
          )}
        </div>
        
        <div className="relative">
          <button 
            onClick={onShowPastConversations}
            onMouseEnter={() => setShowPastConversationsTooltip(true)}
            onMouseLeave={() => setShowPastConversationsTooltip(false)}
            className="flex items-center justify-center"
            aria-label="Past Conversations"
          >
            <ClockIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-white hover:filter hover:brightness-200 transition-all duration-150" />
          </button>
          {showPastConversationsTooltip && (
            <div className="absolute -bottom-8 left-1/2 transform -translate-x-1/2 px-2 py-1 bg-gray-800 text-gray-200 text-xs rounded whitespace-nowrap">
              Past Conversations
            </div>
          )}
        </div>
        
        <BookOpenIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
        <MoreHorizontalIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
        <XIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
      </div>
    </header>
  );
};

export default Header;
