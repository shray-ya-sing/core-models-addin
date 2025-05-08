import * as React from "react";
import { useState } from "react";
import Header from "./Header";
import TailwindFinancialModelChat from "../../client/components/TailwindFinancialModelChat";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FC<AppProps> = () => {
  // Apply global monospace font style
  const monoStyle = {
    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
  };
  
  // State to trigger new conversation
  const [newConversationTrigger, setNewConversationTrigger] = useState(0);
  
  // State to trigger showing past conversations
  const [showPastConversationsTrigger, setShowPastConversationsTrigger] = useState(0);
  
  // State to trigger showing version history
  const [showVersionHistoryTrigger, setShowVersionHistoryTrigger] = useState(0);
  
  // Function to create a new conversation
  const handleNewConversation = () => {
    setNewConversationTrigger(prev => prev + 1);
  };
  
  // Function to show past conversations
  const handleShowPastConversations = () => {
    setShowPastConversationsTrigger(prev => prev + 1);
  };
  
  // Function to show version history
  const handleShowVersionHistory = () => {
    setShowVersionHistoryTrigger(prev => prev + 1);
  };

  return (
    <div className="h-screen flex flex-col bg-transparent" style={monoStyle}>
      <Header 
        onNewConversation={handleNewConversation} 
        onShowPastConversations={handleShowPastConversations}
        onShowVersionHistory={handleShowVersionHistory}
      />
      <div className="flex-1 flex flex-col overflow-auto">
        <TailwindFinancialModelChat 
          newConversationTrigger={newConversationTrigger}
          showPastConversationsTrigger={showPastConversationsTrigger}
          showVersionHistoryTrigger={showVersionHistoryTrigger}
        />
      </div>
    </div>
  );
};

export default App;
