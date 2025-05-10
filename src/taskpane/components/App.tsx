import * as React from "react";
import { useState, useEffect, useRef } from "react";
import Header from "./Header";
import TailwindFinancialModelChat from "../../client/components/TailwindFinancialModelChat";
import StartNewConversation from "../../client/components/StartNewConversation";
import SetupScreen from "../../client/components/SetupScreen";
import { ConversationSession } from "../../client/models/ConversationModels";

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
  // State for test mode visibility
  const [showTestUI, setShowTestUI] = useState(false);
  
  // Toggle test mode visibility
  const handleToggleTestMode = () => {
    setShowTestUI(prev => !prev);
  };
  
  // State to control which component is shown
  const [setupComplete, setSetupComplete] = useState(false);
  const [showChat, setShowChat] = useState(false);
  const [initialMessage, setInitialMessage] = useState('');
  const [resetChat, setResetChat] = useState(true);
  
  // Handle setup completion
  const handleSetupComplete = () => {
    setSetupComplete(true);
  };
  
  // Handle message sent from StartNewConversation
  const handleMessageSent = (message: string) => {
    // Store the message to pass to TailwindFinancialModelChat
    setResetChat(true);
    setInitialMessage(message);
    
    // Switch to chat view
    setShowChat(true);

  };
  
  return (
    <div className="h-screen flex flex-col bg-transparent" style={monoStyle}>
      <Header 
        onToggleTestMode={handleToggleTestMode}
      />
      <div className="flex-1 flex flex-col overflow-auto">
        {!setupComplete ? (
          <SetupScreen onSetupComplete={handleSetupComplete} />
        ) : showChat ? (
          <TailwindFinancialModelChat 
            initialMessage={initialMessage}
            resetChat={resetChat}
          />
        ) : (
          <StartNewConversation 
            onMessageSent={handleMessageSent}
            servicesReady={true}
          />
        )}
      </div>
      
      {/* Conditionally render TestMode only in non-production environments */}
      {process.env.NODE_ENV !== 'production' && TestMode && showTestUI && (
        <React.Suspense fallback={<div style={{ padding: '0.5rem', backgroundColor: '#1a1a1a', color: '#e0e0e0', fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif' }}>Loading test mode...</div>}>
          <TestMode
            queryProcessor={null} // Will be initialized later
            commandInterpreter={null} // Will be initialized later
            onClose={() => setShowTestUI(false)}
          />
        </React.Suspense>
      )}
    </div>
  );
};

export default App;
