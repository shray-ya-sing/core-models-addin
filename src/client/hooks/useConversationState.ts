import { useState, useEffect, useCallback, useRef } from 'react';
import { v4 as uuidv4 } from 'uuid';
import { ChatMessage, ConversationSession } from '../models/ConversationModels';

interface ConversationStateProps {
  currentWorkbookId: string;
}

export const useConversationState = ({ currentWorkbookId }: ConversationStateProps) => {
  const STORAGE_KEY = 'excel-addin-conversations';
  
  // State
  const [currentSession, setCurrentSession] = useState<string>('');
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [sessions, setSessions] = useState<ConversationSession[]>([]);
  const [hasUserSentMessage, setHasUserSentMessage] = useState(false);
  
  // Create a new conversation session
  const createNewSession = useCallback(() => {
    const newSessionId = uuidv4();
    const newSession: ConversationSession = {
      id: newSessionId,
      title: 'New Conversation',
      messages: [],
      lastUpdated: Date.now(),
      createdAt: Date.now(),
      workbookId: currentWorkbookId,
    };

    // Add to sessions
    setSessions((prev) => {
      const updatedSessions = [newSession, ...prev];
      localStorage.setItem(STORAGE_KEY, JSON.stringify(updatedSessions));
      return updatedSessions;
    });

    // Set the current session to the new session
    setCurrentSession(newSessionId);
    
    // Clear the messages for the new session
    setMessages([]);
    
    // Set the flag to indicate we've explicitly created a new session
    // This will prevent automatic session loading from overriding it
    explicitNewSessionRef.current = true;
    
    return newSessionId;
  }, [currentWorkbookId]);
  
  // Function to clear all conversations
  const clearAllConversations = useCallback(() => {
    try {
      localStorage.removeItem(STORAGE_KEY);
      setSessions([]);
      setMessages([]);
      setCurrentSession('');
      setHasUserSentMessage(false);
    } catch (error) {
      console.error('Error clearing conversations:', error);
    }
  }, []);

  // Filter sessions by current workbook ID
  const getWorkbookSessions = useCallback((allSessions: ConversationSession[]) => {
    if (!currentWorkbookId) {
      return allSessions;
    }
    return allSessions.filter((session) => session.workbookId === currentWorkbookId || !session.workbookId);
  }, [currentWorkbookId]);
  
  // Get session title from messages
  const getSessionTitle = useCallback((messages: ChatMessage[]) => {
    const firstUserMessage = messages.find(msg => msg.role === 'user');
    if (firstUserMessage) {
      const truncatedContent = firstUserMessage.content.substring(0, 30);
      return truncatedContent + (firstUserMessage.content.length > 30 ? '...' : '');
    }
    return 'New Conversation';
  }, []);
  
  // Update session messages
  const updateSessionMessages = useCallback((sessionId: string, updatedMessages: ChatMessage[]) => {
    setSessions(prev => {
      return prev.map(session => {
        if (session.id === sessionId) {
          const title = session.title === 'New Conversation' 
            ? getSessionTitle(updatedMessages)
            : session.title;
            
          return {
            ...session,
            messages: updatedMessages,
            title,
            lastUpdated: Date.now()
          };
        }
        return session;
      });
    });
  }, [getSessionTitle]);
  
  // Load a specific session
  const loadSession = useCallback((sessionId: string) => {
    try {
      const session = sessions.find(s => s.id === sessionId);
      if (session) {
        setMessages(session.messages || []);
        setHasUserSentMessage((session.messages && session.messages.length > 0) || false);
        setCurrentSession(sessionId);
      } else {
        console.warn(`Session with ID ${sessionId} not found`);
      }
    } catch (error) {
      console.error('Error loading session:', error);
    }
  }, [sessions]);
  
  // Format time for display
  const formatTimeAgo = useCallback((timestamp: number) => {
    const now = Date.now();
    const secondsAgo = Math.floor((now - timestamp) / 1000);
    
    if (secondsAgo < 60) return `${secondsAgo}s`;
    if (secondsAgo < 3600) return `${Math.floor(secondsAgo / 60)}m`;
    if (secondsAgo < 86400) return `${Math.floor(secondsAgo / 3600)}h`;
    return `${Math.floor(secondsAgo / 86400)}d`;
  }, []);
  
  // Load sessions from localStorage when component mounts
  useEffect(() => {
    const savedSessions = localStorage.getItem(STORAGE_KEY);
    if (savedSessions) {
      try {
        const parsedSessions = JSON.parse(savedSessions) as ConversationSession[];
        setSessions(parsedSessions);
        
        // Load the most recent session
        if (parsedSessions.length > 0) {
          const latestSession = parsedSessions[0];
          setCurrentSession(latestSession.id);
          setMessages(latestSession.messages);
        }
      } catch (error) {
        console.error('Error loading sessions from localStorage:', error);
        localStorage.removeItem(STORAGE_KEY);
      }
    }
  }, []);
  
  // Update localStorage when sessions change
  useEffect(() => {
    if (sessions.length > 0) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    }
  }, [sessions]);
  
  // Update current session's messages when messages state changes
  useEffect(() => {
    if (currentSession && messages.length > 0) {
      updateSessionMessages(currentSession, messages);
    }
  }, [messages, currentSession, updateSessionMessages]);
  
  // Set current session based on workbook ID when it changes
  // We use a ref to track if we've explicitly created a new session to avoid
  // automatic loading overriding it
  const explicitNewSessionRef = useRef<boolean>(false);
  
  useEffect(() => {
    // Skip automatic session loading if we've explicitly created a new session
    if (explicitNewSessionRef.current) {
      console.log('Skipping automatic session loading due to explicit new session');
      explicitNewSessionRef.current = false;
      return;
    }
    
    if (!currentWorkbookId || sessions.length === 0) return;
    
    console.log('Setting session based on workbook ID:', currentWorkbookId);
    
    // Get workbook-specific sessions
    const workbookSessions = getWorkbookSessions(sessions);
    
    // If there are workbook-specific sessions, set the most recent one as current
    if (workbookSessions.length > 0) {
      setCurrentSession(workbookSessions[0].id);
      setMessages(workbookSessions[0].messages);
      console.log('Loaded existing session for workbook:', workbookSessions[0].title);
    } else {
      // If no workbook-specific sessions, create a new one
      console.log('No sessions for current workbook, creating new session');
      createNewSession();
    }
  }, [currentWorkbookId, sessions.length, getWorkbookSessions, createNewSession]);
  
  return {
    currentSession,
    setCurrentSession,
    messages,
    setMessages,
    sessions,
    setSessions,
    hasUserSentMessage,
    setHasUserSentMessage,
    createNewSession,
    clearAllConversations,
    getWorkbookSessions,
    updateSessionMessages,
    loadSession,
    formatTimeAgo,
    // Export the explicit new session flag for external use if needed
    setExplicitNewSession: (value: boolean) => { explicitNewSessionRef.current = value; }
  };
};

export default useConversationState;