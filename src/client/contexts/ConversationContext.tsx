import React, { createContext, useContext, ReactNode } from 'react';

interface ConversationContextType {
  createNewSession: () => void;
}

const ConversationContext = createContext<ConversationContextType | undefined>(undefined);

export const ConversationProvider: React.FC<{
  children: ReactNode;
  createNewSession: () => void;
}> = ({ children, createNewSession }) => {
  return (
    <ConversationContext.Provider value={{ createNewSession }}>
      {children}
    </ConversationContext.Provider>
  );
};

export const useConversation = (): ConversationContextType => {
  const context = useContext(ConversationContext);
  if (context === undefined) {
    throw new Error('useConversation must be used within a ConversationProvider');
  }
  return context;
};
