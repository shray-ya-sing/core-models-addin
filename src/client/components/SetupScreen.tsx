import * as React from 'react';
import { useState, useEffect } from 'react';
import { LoadContextService } from '../services/context/LoadContextService';
import { ClientAnthropicService } from '../services/llm/ClientAnthropicService';
import { ClientWorkbookStateManager } from '../services/context/ClientWorkbookStateManager';
import { ClientSpreadsheetCompressor } from '../services/context/ClientSpreadsheetCompressor';

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
  const [contextService, setContextService] = useState<LoadContextService | null>(null);
  
  // Initialize services and build context
  useEffect(() => {
    const initializeServices = async () => {
      try {
        // Step 1: Initialize base services
        setMessage('Initializing services...');
        setProgress(10);
        
        // Use environment variable or config for API key in production
        const anthropicService = new ClientAnthropicService(process.env.ANTHROPIC_API_KEY || '');
        const workbookStateManager = new ClientWorkbookStateManager(5000); // 5 second cache timeout
        // ClientSpreadsheetCompressor doesn't require constructor parameters
        const compressor = new ClientSpreadsheetCompressor();
        
        setProgress(30);
        setMessage('Setting up context service...');
        
        // Step 2: Get or create the LoadContextService singleton
        const loadContextService = LoadContextService.getInstance({
          workbookStateManager: workbookStateManager,
          compressor: compressor,
          anthropic: anthropicService,
          useAdvancedChunkLocation: true
        });
        
        setContextService(loadContextService);
        
        // Step 3: Check if we have cached metadata for this workbook
        setProgress(40);
        setMessage('Checking for cached metadata...');
        
        const workbookId = workbookStateManager.getWorkbookId();
        const hasCachedMetadata = loadContextService.hasMetadataForCurrentWorkbook();
        
        if (hasCachedMetadata) {
          console.log(`Using cached metadata for workbook: ${workbookId}`);
          setMessage(`Loading workbook data`);
          setProgress(100);
          setMessage('Successfully loaded workbook data');
          
          // Transition to StartNewConversation after a short delay
          setTimeout(() => {
            if (onSetupComplete) {
              onSetupComplete();
            }
          }, 800); // Short delay for user to see the success message
        } else {
          console.log(`Building new metadata for workbook: ${workbookId}`);
          setMessage(`Loading workbook data`);
          // Only force cache refresh during initial setup
          loadContextService.setupCache(true);          
          setProgress(100);
          setMessage('Successfully loaded workbook data');
          
          // Transition to StartNewConversation after a short delay
          setTimeout(() => {
            if (onSetupComplete) {
              onSetupComplete();
            }
          }, 800); // Short delay for user to see the success message
        }
        
      } catch (error) {
        console.error('Error initializing services:', error);
        setMessage('Error initializing services. Please refresh the page.');
      }
    };
    
    initializeServices();
  }, [onSetupComplete]);
  return (
    <div className="flex flex-col h-full items-center justify-center p-4 bg-gray-900">
      {/* Logo */}
      <div className="mb-6">
        <img 
          src="assets/cori-logo.svg" 
          style={{ 
            width: '48px', 
            height: '48px', 
            opacity: 0.9,
            fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif'
          }} 
          alt="Cori Logo" 
        />  
      </div>
      
      {/* Message */}
      <p 
        className="text-white/80 text-center mb-6" 
        style={{ 
          fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
          fontSize: '14px'
        }}
      >
        {message}
      </p>
      
      {/* Progress bar */}
      <div 
        className="w-64 bg-gray-800 rounded-full h-2.5 mb-4 overflow-hidden"
        style={{ border: '1px solid rgba(255, 255, 255, 0.1)' }}
      >
        <div 
          className="bg-blue-500 h-2.5 rounded-full transition-all duration-300 ease-out"
          style={{ width: `${progress}%` }}
        ></div>
      </div>
      
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
    </div>
  );
};

export default SetupScreen;
