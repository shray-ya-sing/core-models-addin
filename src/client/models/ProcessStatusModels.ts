/**
 * Models for tracking process status throughout the application
 */

export enum ProcessStage {
  KnowledgeBaseQuery = 'knowledge_base_query',
  WorkbookCapture = 'workbook_capture',
  QueryProcessing = 'query_processing',
  ResponseGeneration = 'response_generation',
  CommandPlanning = 'command_planning',
  CommandExecution = 'command_execution',
  OperationExecution = 'operation_execution'
}

export enum ProcessStatus {
  Pending = 'pending',
  Success = 'success',
  Error = 'error',
  Idle = 'idle'
}

export interface ProcessStatusEvent {
  id: string;
  stage: ProcessStage;
  status: ProcessStatus;
  message: string;
  timestamp: number;
  error?: Error;
  data?: any;
}

export interface ProcessStatusUpdate {
  stage: ProcessStage;
  status: ProcessStatus;
  message: string;
  error?: Error;
  data?: any;
}

export type ProcessStatusListener = (event: ProcessStatusEvent) => void;

/**
 * Process Manager for tracking statuses across different parts of the app
 */
export class ProcessStatusManager {
  private static instance: ProcessStatusManager;
  private listeners: ProcessStatusListener[] = [];
  private statuses: Map<string, ProcessStatusEvent> = new Map();

  private constructor() {}

  /**
   * Get the singleton instance
   */
  public static getInstance(): ProcessStatusManager {
    if (!ProcessStatusManager.instance) {
      ProcessStatusManager.instance = new ProcessStatusManager();
    }
    return ProcessStatusManager.instance;
  }

  /**
   * Add a listener for process status updates
   * @param listener The event listener
   * @returns Function to remove the listener
   */
  public addListener(listener: ProcessStatusListener): () => void {
    this.listeners.push(listener);
    return () => {
      this.listeners = this.listeners.filter(l => l !== listener);
    };
  }

  /**
   * Update the status of a process stage
   * @param processId The ID of the process (e.g., query ID or command ID)
   * @param update The status update
   */
  public updateStatus(processId: string, update: ProcessStatusUpdate): void {
    const event: ProcessStatusEvent = {
      id: processId,
      stage: update.stage,
      status: update.status,
      message: update.message,
      timestamp: Date.now(),
      error: update.error,
      data: update.data
    };

    // Store status
    this.statuses.set(`${processId}_${update.stage}`, event);

    // Notify listeners
    this.listeners.forEach(listener => listener(event));
  }

  /**
   * Get all statuses for a specific process
   * @param processId The ID of the process
   * @returns Array of status events
   */
  public getStatusesForProcess(processId: string): ProcessStatusEvent[] {
    const processStatuses: ProcessStatusEvent[] = [];
    
    this.statuses.forEach((status) => {
      if (status.id === processId) {
        processStatuses.push(status);
      }
    });
    
    return processStatuses.sort((a, b) => a.timestamp - b.timestamp);
  }

  /**
   * Clear all statuses
   */
  public clearAllStatuses(): void {
    this.statuses.clear();
  }

  /**
   * Get the status of a specific process
   * @param processId The ID of the process
   * @returns The status event for the process
   */
  public getProcess(processId: string): ProcessStatusEvent | undefined {
    return this.statuses.get(processId);
  }
}
