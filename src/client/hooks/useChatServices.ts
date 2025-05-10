// src/client/hooks/useChatServices.ts
import { useState, useEffect, useCallback } from 'react';
import { v4 as uuidv4 } from 'uuid';
import { ClientCommandManager } from '../services/actions/ClientCommandManager';
import { ClientWorkbookStateManager } from '../services/context/ClientWorkbookStateManager';
import { ClientSpreadsheetCompressor } from '../services/context/ClientSpreadsheetCompressor';
import { ClientExcelCommandAdapter } from '../services/actions/ClientExcelCommandAdapter';
import { ClientExcelCommandInterpreter } from '../services/actions/ClientExcelCommandInterpreter';
import { ClientAnthropicService } from '../services/llm/ClientAnthropicService';
import { ClientKnowledgeBaseService } from '../services/ClientKnowledgeBaseService';
import { ClientQueryProcessor } from '../services/request-processing/ClientQueryProcessor';
import { OpenAIClientService } from '../services/llm/OpenAIClientService';
import { initializeMultimodalAnalysisService } from '../services/document-understanding/MultimodalAnalysisService';
import { VersionHistoryProvider } from '../services/versioning/VersionHistoryProvider';
import { AIApprovalSystem } from '../services/pending-changes/AIApprovalSystem';
import { PendingChangesTracker } from '../services/pending-changes/PendingChangesTracker';
import { ShapeEventHandler } from '../services/ShapeEventHandler';
import { CommandStatus } from '../models/CommandModels';
import { ProcessStatusManager, ProcessStatus } from '../models/ProcessStatusModels';
import { StatusType } from '../components/StatusIndicator';
import { LoadContextService } from '../services/context/LoadContextService';  
import config from '../config';

export interface ChatServices {
  commandManager: ClientCommandManager | null;
  loadContextService: LoadContextService | null;
  workbookStateManager: ClientWorkbookStateManager | null;
  spreadsheetCompressor: ClientSpreadsheetCompressor | null;
  anthropicService: ClientAnthropicService | null;
  knowledgeBaseService: ClientKnowledgeBaseService | null;
  queryProcessor: ClientQueryProcessor | null;
  commandInterpreter: ClientExcelCommandInterpreter | null;
  versionHistoryProvider: VersionHistoryProvider;
  pendingChangesTracker: PendingChangesTracker | null;
  shapeEventHandler: ShapeEventHandler | null;
  servicesReady: boolean;
  currentWorkbookId: string;
  approvalEnabled: boolean;
  setApprovalEnabled: (enabled: boolean) => void;
  pendingChanges: any[];
  refreshPendingChanges: () => void;
  handleAcceptAll: () => Promise<void>;
  handleRejectAll: () => Promise<void>;
  handleAcceptChange: (changeId: string) => Promise<void>;
  handleRejectChange: (changeId: string) => Promise<void>;
}

export const useChatServices = (
): ChatServices => {
  // Service instances
  const [commandManager, setCommandManager] = useState<ClientCommandManager | null>(null);
  const [loadContextService, setLoadContextService] = useState<LoadContextService | null>(null);
  const [workbookStateManager, setWorkbookStateManager] = useState<ClientWorkbookStateManager | null>(null);
  const [spreadsheetCompressor, setSpreadsheetCompressor] = useState<ClientSpreadsheetCompressor | null>(null);
  const [anthropicService, setAnthropicService] = useState<ClientAnthropicService | null>(null);
  const [knowledgeBaseService, setKnowledgeBaseService] = useState<ClientKnowledgeBaseService | null>(null);
  const [queryProcessor, setQueryProcessor] = useState<ClientQueryProcessor | null>(null);
  const [commandInterpreter, setCommandInterpreter] = useState<ClientExcelCommandInterpreter | null>(null);
  const [versionHistoryProvider] = useState<VersionHistoryProvider>(() => new VersionHistoryProvider());
  const [pendingChangesTracker, setPendingChangesTracker] = useState<PendingChangesTracker | null>(null);
  const [shapeEventHandler, setShapeEventHandler] = useState<ShapeEventHandler | null>(null);
  
  // Status state
  const [servicesReady, setServicesReady] = useState(false);
  const [currentWorkbookId, setCurrentWorkbookId] = useState<string>('');
  const [approvalEnabled, setApprovalEnabled] = useState<boolean>(false);
  const [pendingChanges, setPendingChanges] = useState<any[]>([]);
  
  // Get current workbook ID
  const getCurrentWorkbookId = async (): Promise<string> => {
    try {
      if (typeof Excel === 'undefined') {
        console.warn('Excel API not available');
        return 'unknown-workbook';
      }
      
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load('name');
        
        await context.sync();
        
        const workbookId = workbook.name || `workbook-${new Date().getTime()}`;
        return workbookId;
      });
    } catch (error) {
      console.error('Error getting workbook ID:', error);
      return 'unknown-workbook';
    }
  };
  
  // Function to refresh pending changes
  const refreshPendingChanges = useCallback(() => {
    if (pendingChangesTracker && currentWorkbookId) {
      console.log('Refreshing pending changes for workbook:', currentWorkbookId);
      const changes = pendingChangesTracker.getPendingChanges(currentWorkbookId);
      console.log('Found pending changes:', changes.length);
      setPendingChanges(changes);
    }
  }, [pendingChangesTracker, currentWorkbookId]);
  
  // Functions to handle accept/reject actions
  const handleAcceptAll = useCallback(async () => {
    if (pendingChangesTracker && pendingChanges.length > 0) {
      for (const change of pendingChanges) {
        await pendingChangesTracker.acceptChange(change.id);
      }
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, pendingChanges, refreshPendingChanges]);
  
  const handleRejectAll = useCallback(async () => {
    if (pendingChangesTracker && pendingChanges.length > 0) {
      for (const change of pendingChanges) {
        await pendingChangesTracker.rejectChange(change.id);
      }
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, pendingChanges, refreshPendingChanges]);
  
  const handleAcceptChange = useCallback(async (changeId: string) => {
    if (pendingChangesTracker) {
      await pendingChangesTracker.acceptChange(changeId);
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, refreshPendingChanges]);
  
  const handleRejectChange = useCallback(async (changeId: string) => {
    if (pendingChangesTracker) {
      await pendingChangesTracker.rejectChange(changeId);
      refreshPendingChanges();
    }
  }, [pendingChangesTracker, refreshPendingChanges]);
  
  // Initialize services
  useEffect(() => {
    const initializeClientServices = async () => {
      const workbookId = await getCurrentWorkbookId();
      setCurrentWorkbookId(workbookId);
      console.log('Current workbook ID:', workbookId);

      try {
        console.log('%c Initializing client services...', 'background: #222; color: #bada55; font-size: 14px');
        
        if (!config.anthropicApiKey) {
          console.warn('No Anthropic API key found in configuration');
        }
        
        // Create service instances
        const interpreter = new ClientExcelCommandInterpreter();
        
        // Initialize version history system
        versionHistoryProvider.initialize(interpreter);
        versionHistoryProvider.setCurrentWorkbookId(workbookId);
        
        // Create other services
        const stateManager = new ClientWorkbookStateManager();
        const compressor = new ClientSpreadsheetCompressor();
        const adapter = new ClientExcelCommandAdapter(interpreter);
        const manager = new ClientCommandManager(stateManager, adapter);
        
        const anthropic = new ClientAnthropicService(config.anthropicApiKey, config.openaiApiKey);
        const knowledgeBase = new ClientKnowledgeBaseService(config.knowledgeBaseApiUrl);
        // Use the existing LoadContextService singleton or create a new one if it doesn't exist
        const loadContextService = LoadContextService.getInstance({
          workbookStateManager: stateManager,
          compressor,
          anthropic,
          useAdvancedChunkLocation: true
        });
        
        // Initialize OpenAI client service for fallback
        const openai = new OpenAIClientService(config.openaiApiKey);
        
        const processor = new ClientQueryProcessor({
          anthropic,
          kbService: knowledgeBase,
          contextService: loadContextService,
          commandManager: manager,
          openai,
          useOpenAIFallback: true
        });
        
        // Set up event listeners for workbook changes
        await stateManager.setupChangeListeners();
        
        // Set workbook ID in version history provider
        versionHistoryProvider.setCurrentWorkbookId(workbookId);
        
        // Initialize AI Approval System
        const versionHistoryService = versionHistoryProvider.getVersionHistoryService();
        const { pendingChangesTracker: pct, shapeEventHandler: seh } = AIApprovalSystem.initialize(interpreter, versionHistoryService);
        
        // Set workbook ID on shape event handler and start polling
        seh.setCurrentWorkbookId(workbookId);
        seh.startPolling();
        
        // Update state with service instances
        setWorkbookStateManager(stateManager);
        setSpreadsheetCompressor(compressor);
        setCommandManager(manager);
        setAnthropicService(anthropic);
        setKnowledgeBaseService(knowledgeBase);
        setQueryProcessor(processor);
        setCommandInterpreter(interpreter);
        setPendingChangesTracker(pct);
        setShapeEventHandler(seh);
        
        // Enable approval workflow by default
        interpreter.setRequireApproval(true);
        setApprovalEnabled(true);
        
        // Register for command updates
        const unsubscribeCommandUpdate = manager.onCommandUpdate((command) => {
          console.log('Command update received:', command);
          
          if (command.status === CommandStatus.Completed) {
            console.log(`Command "${command.description}" completed successfully.`);
            
            // Refresh pending changes after command execution
            setTimeout(() => {
              if (pct && workbookId) {
                const changes = pct.getPendingChanges(workbookId);
                setPendingChanges(changes);
              }
            }, 500);
          } else if (command.status === CommandStatus.Failed) {
            console.error(`Command "${command.description}" failed: ${command.error || 'Unknown error'}`);
          }
        });
        
        // Set up process status manager listener
        const processManager = ProcessStatusManager.getInstance();
        const unsubscribeProcessEvents = processManager.addListener((event) => {
          console.log('Process Status Event:', event);
          
          // Map process status to StatusIndicator status
          let statusType = StatusType.Idle;
          switch (event.status) {
            case ProcessStatus.Pending:
              statusType = StatusType.Pending;
              break;
            case ProcessStatus.Success:
              statusType = StatusType.Success;
              break;
            case ProcessStatus.Error:
              statusType = StatusType.Error;
              break;
          }
          
          // Create the status message
          const statusMessage = {
            role: 'status' as const,
            content: event.message,
            status: statusType,
            stage: event.stage
          };

        });
        
        // Mark services as ready
        console.log('%c All services initialized successfully!', 'background: #222; color: #2ecc71; font-size: 14px');
        setServicesReady(true);
        
        // Return cleanup function
        return () => {
          if (unsubscribeCommandUpdate) unsubscribeCommandUpdate();
          if (unsubscribeProcessEvents) unsubscribeProcessEvents();
          
          if (seh) {
            seh.stopPolling();
          }
        };
      } catch (error) {
        console.error('Error initializing client services:', error);
        return () => {}; // Return empty cleanup function for error case
      }
    };
    
    const cleanup = initializeClientServices();
    
    return () => {
      if (cleanup) {
        cleanup.then(cleanupFn => {
          if (cleanupFn) cleanupFn();
        }).catch(err => {
          console.error('Error in cleanup function:', err);
        });
      }
    };
  }, []);

  return {
    commandManager,
    workbookStateManager,
    spreadsheetCompressor,
    loadContextService,
    anthropicService,
    knowledgeBaseService,
    queryProcessor,
    commandInterpreter,
    versionHistoryProvider,
    pendingChangesTracker,
    shapeEventHandler,
    servicesReady,
    currentWorkbookId,
    approvalEnabled,
    setApprovalEnabled,
    pendingChanges,
    refreshPendingChanges,
    handleAcceptAll,
    handleRejectAll,
    handleAcceptChange,
    handleRejectChange
  };
};