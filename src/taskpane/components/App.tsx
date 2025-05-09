import * as React from "react";
import { useState, useEffect } from "react";
import Header from "./Header";
import TailwindFinancialModelChat from "../../client/components/TailwindFinancialModelChat";

// Conditionally import TestMode only in non-production environments
const TestMode = process.env.NODE_ENV !== 'production'
  ? React.lazy(() => import('../../test-ui/TestMode'))
  : null;

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FC<AppProps> = () => {
  // Apply global Arial font style
  const monoStyle = {
    fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
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

  // State for query processor and command interpreter
  const [queryProcessor, setQueryProcessor] = useState(null);
  const [commandInterpreter, setCommandInterpreter] = useState(null);
  
  // State for test mode visibility
  const [showTestUI, setShowTestUI] = useState(false);
  
  // Get references to the query processor and command interpreter from the chat component
  const handleComponentsReady = (components) => {
    if (components.queryProcessor && components.commandInterpreter) {
      setQueryProcessor(components.queryProcessor);
      setCommandInterpreter(components.commandInterpreter);
    }
  };
  
  // Toggle test mode visibility
  const handleToggleTestMode = () => {
    setShowTestUI(prev => !prev);
  };
  
  return (
    <div className="h-screen flex flex-col bg-transparent" style={monoStyle}>
      <Header 
        onNewConversation={handleNewConversation} 
        onShowPastConversations={handleShowPastConversations}
        onShowVersionHistory={handleShowVersionHistory}
        onToggleTestMode={handleToggleTestMode}
      />
      <div className="flex-1 flex flex-col overflow-auto">
        <TailwindFinancialModelChat 
          newConversationTrigger={newConversationTrigger}
          showPastConversationsTrigger={showPastConversationsTrigger}
          showVersionHistoryTrigger={showVersionHistoryTrigger}
          onComponentsReady={handleComponentsReady}
        />
      </div>
      
      {/* Conditionally render TestMode only in non-production environments */}
      {process.env.NODE_ENV !== 'production' && TestMode && queryProcessor && commandInterpreter && showTestUI && (
        <React.Suspense fallback={<div style={{ padding: '0.5rem', backgroundColor: '#1a1a1a', color: '#e0e0e0', fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif' }}>Loading test mode...</div>}>
          <TestMode
            queryProcessor={queryProcessor}
            commandInterpreter={commandInterpreter}
            onClose={() => setShowTestUI(false)}
          />
        </React.Suspense>
      )}
    </div>
  );
};

export default App;
