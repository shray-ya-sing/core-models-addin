import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { LoadContextService } from '../services/context/LoadContextService';
import { ClientAnthropicService } from '../services/llm/ClientAnthropicService';
import { ClientWorkbookStateManager } from '../services/context/ClientWorkbookStateManager';
import { ClientSpreadsheetCompressor } from '../services/context/ClientSpreadsheetCompressor';
import { multimodalAnalysisService, initializeMultimodalAnalysisService } from '../services/document-understanding/MultimodalAnalysisService';

interface SetupScreenProps {
  progress?: number;
  message?: string;
  onSetupComplete?: () => void;
}

/**
 * A setup screen that initializes services and builds context
 */
const SetupScreen: React.FC<SetupScreenProps> = ({ 
  progress: initialProgress = 0, 
  message: initialMessage = 'Setting up...', 
  onSetupComplete 
}) => {
  // State for tracking progress and messages
  const [progress, setProgress] = useState(initialProgress);
  const [message, setMessage] = useState(initialMessage);
  const [displayMessage, setDisplayMessage] = useState('');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  
  // Messages for typewriter effect
  const [messages, setMessages] = useState<string[]>([
    'Initializing services...',
    'Gathering internal resources...',
    'Retrieving workbook context',
    'Understanding links and relationships...',
    'Looking through worksheets...',
    'Analyzing workbook formatting',
    'Loading context data...',
    'Setting up final configurations...',
    'Connecting to internal services...',
    'Analyzing sheet content...',
  ]);
  const [currentMessageIndex, setCurrentMessageIndex] = useState(0);
  const [charIndex, setCharIndex] = useState(0);
  
  // Refs for animation control
  const typewriterTimerRef = useRef<NodeJS.Timeout | null>(null);
  const messageChangeTimerRef = useRef<NodeJS.Timeout | null>(null);

  // Typewriter effect implementation
  useEffect(() => {
    // Clear any existing timers
    if (typewriterTimerRef.current) {
      clearTimeout(typewriterTimerRef.current);
    }
    
    if (charIndex < messages[currentMessageIndex].length) {
      // Continue typing the current message
      typewriterTimerRef.current = setTimeout(() => {
        setDisplayMessage(messages[currentMessageIndex].substring(0, charIndex + 1));
        setCharIndex(charIndex + 1);
      }, 50); // Speed of typing
    } else {
      // Finished typing this message, wait then move to next message
      messageChangeTimerRef.current = setTimeout(() => {
        setCurrentMessageIndex((prevIndex) => (prevIndex + 1) % messages.length);
        setCharIndex(0);
      }, 2000); // Pause after completing message
    }
    
    // Cleanup function
    return () => {
      if (typewriterTimerRef.current) {
        clearTimeout(typewriterTimerRef.current);
      }
      if (messageChangeTimerRef.current) {
        clearTimeout(messageChangeTimerRef.current);
      }
    };
  }, [charIndex, currentMessageIndex, messages]);
  
  // Update the messages array when progress updates
  useEffect(() => {
    if (message && !messages.includes(message)) {
      setMessages(prevMessages => [...prevMessages, message]);
    }
  }, [message, messages]);
  
  // Initialize services on component mount
  useEffect(() => {
    const initializeSetup = async () => {
      try {
        // Step 1: Initialize base services
        setMessage('Initializing services...');
        setProgress(10);
        
        // Create service instances
        const workbookStateManager = new ClientWorkbookStateManager(5000); // 5 second cache timeout
        const anthropicService = new ClientAnthropicService(process.env.ANTHROPIC_API_KEY || '');
        const compressor = new ClientSpreadsheetCompressor();
        
        // Initialize the multimodal analysis service with the Anthropic service
        initializeMultimodalAnalysisService(anthropicService);
        
        setProgress(20);
        setMessage('Setting up context service...');
        
        // Get or create the LoadContextService singleton
        const contextService = LoadContextService.getInstance({
          workbookStateManager,
          compressor,
          anthropic: anthropicService,
          useAdvancedChunkLocation: true
        });
        
        setProgress(30);
        setMessage('Resolving workbook ID...');
        
        // Step 2: Get a valid workbook ID with retries
        let workbookId = workbookStateManager.getWorkbookId();
        
        // If no workbook ID exists, try to get one with retries
        if (!workbookId || workbookId === 'null' || workbookId === 'undefined') {
          console.log('\n=======================================================');
          console.log('🔄 ATTEMPTING TO GET VALID WORKBOOK ID');
          console.log('=======================================================');
          
          // Try to force set a new workbook ID with retries
          const maxRetries = 5;
          const retryDelayMs = 500;
          let retryCount = 0;
          
          while (retryCount < maxRetries && (!workbookId || workbookId === 'null' || workbookId === 'undefined')) {
            try {
              console.log(`Retry ${retryCount + 1}/${maxRetries}: Attempting to get workbook ID...`);
              
              // Get current workbook ID in case it was set elsewhere
              workbookId = workbookStateManager.getWorkbookId();
              
              if (workbookId && workbookId !== 'null' && workbookId !== 'undefined') {
                console.log(`✅ Found valid workbook ID: ${workbookId}`);
                break;
              }
              
              // Try to get a temporary workbook ID by creating one based on timestamp
              const tempId = `excel-workbook-${Date.now()}`;
              workbookStateManager.storeWorkbookId(tempId);
              workbookId = tempId;
              console.log(`✅ Created temporary workbook ID: ${workbookId}`);
              
              retryCount++;
              if (retryCount < maxRetries && (!workbookId || workbookId === 'null' || workbookId === 'undefined')) {
                console.log(`⚠️ No valid workbook ID yet, waiting ${retryDelayMs}ms before retry ${retryCount + 1}/${maxRetries}`);
                await new Promise(resolve => setTimeout(resolve, retryDelayMs));
              }
            } catch (error) {
              console.warn(`⚠️ Error in workbook ID resolution (retry ${retryCount + 1}/${maxRetries}):`, error);
              retryCount++;
              await new Promise(resolve => setTimeout(resolve, retryDelayMs));
            }
          }
        }
        
        if (!workbookId || workbookId === 'null' || workbookId === 'undefined') {
          console.warn('\n=======================================================');
          console.warn('⚠️ FAILED TO GET VALID WORKBOOK ID AFTER RETRIES');
          console.warn('=======================================================');
          setMessage('Warning: Unable to identify workbook');
          setProgress(40);
          // Even without a workbook ID, try to continue initialization
          workbookId = 'default-workbook-id';
          workbookStateManager.storeWorkbookId(workbookId);
        } else {
          console.log(`\n=======================================================`);
          console.log(`✅ OBTAINED VALID WORKBOOK ID: ${workbookId}`);
          console.log(`=======================================================`);
        }
        
        // Make sure the ID is stored in the state manager
        if (workbookId) {
          workbookStateManager.storeWorkbookId(workbookId);
        }
        
        // Step 4: Check if we have cached metadata for this workbook
        const hasCachedMetadata = contextService.hasMetadataForCurrentWorkbook();
        console.log(`Has cached metadata: ${hasCachedMetadata ? 'YES' : 'NO'}`);
        
        // Step 5: Analyze formatting protocol only if we have a valid workbook ID
        setMessage('Analyzing workbook formatting protocol...');
        setProgress(50);
        
        try {
          console.log('\n=======================================================');
          console.log('🔍 STARTING FORMATTING PROTOCOL ANALYSIS AND WAITING FOR COMPLETION');
          console.log(`🔍 USING WORKBOOK ID: ${workbookId}`);
          console.log('=======================================================');
          
          // Make sure the multimodal analysis service has the correct workbook ID
          // and wait for the formatting analysis to complete (blocking)
          await multimodalAnalysisService.setWorkbookAndEnsureFormatting(workbookId, true);
          
          // Verify the formatting protocol was created and cached
          const formattingProtocol = multimodalAnalysisService.getWorkbookFormattingProtocol(workbookId);
          
          if (formattingProtocol && Object.keys(formattingProtocol).length > 0) {
            console.log('\n=======================================================');
            console.log('✅ FORMATTING PROTOCOL ANALYSIS COMPLETED SUCCESSFULLY');
            console.log(`Protocol contains ${Object.keys(formattingProtocol).length} categories: ${Object.keys(formattingProtocol).join(', ')}`);
            console.log('=======================================================');
            
            setProgress(70);
            setMessage('Formatting protocol analysis completed successfully');
          } else {
            console.warn('\n=======================================================');
            console.warn('⚠️ FORMATTING PROTOCOL ANALYSIS COMPLETED BUT NO VALID PROTOCOL WAS CACHED');
            console.warn('Using default formatting protocol instead');
            console.warn('=======================================================');
            
            setProgress(60);
            setMessage('Using default formatting protocol');
          }
        } catch (error) {
          console.warn('\n=======================================================');
          console.warn('⚠️ Error initializing formatting protocol:', error);
          console.warn('=======================================================');
          setMessage('Continuing with default formatting protocol...');
        }
        
        // Step 6: Load context data
        setMessage('Loading workbook context data...');
        const forceCacheRefresh = !hasCachedMetadata;
        contextService.setupCache(forceCacheRefresh);
        
        setProgress(90);
        setMessage('Workbook data loaded successfully');
        
        // Step 7: Complete setup and transition to conversation
        setProgress(100);
        
        console.log('\n=======================================================');
        console.log('🚀 SETUP COMPLETE - TRANSITIONING TO CONVERSATION COMPONENT');
        console.log('=======================================================');
        
        // Transition to the conversation component after a short delay
        setTimeout(() => {
          if (onSetupComplete) {
            onSetupComplete();
          }
        }, 1000);
      } catch (error) {
        console.error('Error in setup screen initialization:', error);
        setErrorMessage(`Setup failed: ${error.message || 'Unknown error'}`);
        setProgress(0);
        setMessage('Setup failed. Please try again.');
      }
    };
    
    initializeSetup();
  }, [onSetupComplete]);  // Only run once on mount, but include onSetupComplete in dependencies

  return (
    <div className="flex flex-col h-full items-center justify-center p-4 bg-gray-900">
      {/* Title */}
      <h1 
        className="text-2xl text-white font-semibold mb-8" 
        style={{ 
          fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
          letterSpacing: '0.025em'
        }}
      >
        Cori is setting up
      </h1>
      
      {/* Error message (if any) */}
      {errorMessage && (
        <div className="bg-red-900/50 text-white p-3 rounded mb-4 w-64 text-center">
          {errorMessage}
        </div>
      )}
      
      {/* Spinner */}
      <div className="mb-6">
        <div className="spinner">
          <div className="spinner-circle"></div>
          <div className="spinner-circle"></div>
          <div className="spinner-circle"></div>
        </div>
      </div>
      
      {/* Status message with typewriter effect */}
      <p 
        className="text-white/70 text-center mb-6 min-h-[3rem]" 
        style={{ 
          fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
          fontSize: '14px',
          maxWidth: '300px'
        }}
      >
        {displayMessage}<span className="text-cursor">|</span>
      </p>
      
      {/* Progress text */}
      <p 
        className="text-white/50 text-center" 
        style={{ 
          fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
          fontSize: '12px'
        }}
      >
        {progress > 0 ? `${Math.round(progress)}%` : 'Initializing...'}
      </p>
      
      {/* Adding animation styles to head on component mount */}
      <style dangerouslySetInnerHTML={{ __html: `
        .spinner {
          display: flex;
          justify-content: center;
          align-items: center;
          gap: 8px;
        }
        
        .spinner-circle {
          width: 12px;
          height: 12px;
          background-color: #3b82f6;
          border-radius: 50%;
          animation: bounce 1.4s infinite ease-in-out both;
        }
        
        .spinner-circle:nth-child(1) {
          animation-delay: -0.32s;
        }
        
        .spinner-circle:nth-child(2) {
          animation-delay: -0.16s;
        }
        
        @keyframes bounce {
          0%, 80%, 100% { 
            transform: scale(0);
          } 40% { 
            transform: scale(1.0);
          }
        }
        
        .text-cursor {
          animation: blink 1s step-end infinite;
        }
        
        @keyframes blink {
          from, to { opacity: 1; }
          50% { opacity: 0; }
        }
      `}} />
    </div>
  );
};

export default SetupScreen;