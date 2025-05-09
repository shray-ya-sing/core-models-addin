import { Command, CommandStatus, CommandStep, Operation } from '../models/CommandModels';
import { ClientCommandExecutor } from './ClientCommandExecutor';
import { ClientExcelCommandAdapter } from './ClientExcelCommandAdapter';
import { ClientWorkbookStateManager } from './ClientWorkbookStateManager';

/**
 * Client-side command manager for tracking and executing commands
 */
export class ClientCommandManager {
  private commands: Map<string, Command> = new Map();
  private commandExecutor: ClientCommandExecutor;
  private excelCommandAdapter: ClientExcelCommandAdapter;
  private workbookStateManager: ClientWorkbookStateManager | null = null;
  private commandUpdateListeners: ((command: Command) => void)[] = [];

  /**
   * Create a new ClientCommandManager
   * @param commandExecutor The command executor to use
   * @param workbookStateManager Optional workbook state manager
   * @param excelCommandAdapter Optional existing adapter instance to use. If not provided, a new one will be created.
   */
  constructor(
    commandExecutor: ClientCommandExecutor, 
    workbookStateManager?: ClientWorkbookStateManager,
    excelCommandAdapter?: ClientExcelCommandAdapter
  ) {
    this.commandExecutor = commandExecutor;
    this.excelCommandAdapter = excelCommandAdapter || new ClientExcelCommandAdapter();
    this.workbookStateManager = workbookStateManager || null;
    console.log(`🔄 [ClientCommandManager] Using ${excelCommandAdapter ? 'provided' : 'new'} adapter instance`);
  }

  /**
   * Add a command to the manager
   * @param command The command to add
   */
  public addCommand(command: Command): void {
    // Set initial state if not already set
    if (!command.createdAt) {
      command.createdAt = new Date();
    }
    if (!command.updatedAt) {
      command.updatedAt = new Date();
    }
    if (!command.status) {
      command.status = CommandStatus.Pending;
    }
    if (!command.progress) {
      command.progress = 0;
    }

    // Initialize step statuses
    command.steps.forEach(step => {
      if (!step.status) {
        step.status = 'pending';
      }
      
      // Initialize operation statuses
      step.operations.forEach(operation => {
        if (!operation.status) {
          operation.status = 'pending';
        }
      });
    });

    this.commands.set(command.id, command);
    this.notifyListeners(command);
  }

  /**
   * Get a command by ID
   * @param commandId The command ID
   * @returns The command or undefined if not found
   */
  public getCommand(commandId: string): Command | undefined {
    return this.commands.get(commandId);
  }

  /**
   * Get all commands
   * @returns Array of all commands
   */
  public getAllCommands(): Command[] {
    return Array.from(this.commands.values());
  }

  /**
   * Update a command's status
   * @param commandId The command ID
   * @param status The new status
   * @param error Optional error message
   */
  public updateCommandStatus(commandId: string, status: CommandStatus, error?: string): void {
    const command = this.commands.get(commandId);
    if (command) {
      command.status = status;
      command.updatedAt = new Date();
      
      if (error) {
        command.error = error;
      }
      
      // Update progress based on status
      if (status === CommandStatus.Completed) {
        command.progress = 100;
      } else if (status === CommandStatus.Failed) {
        // Keep progress as is
      }
      
      this.notifyListeners(command);
    }
  }

  /**
   * Update a step's status
   * @param commandId The command ID
   * @param stepIndex The step index
   * @param status The new status
   * @param error Optional error message
   */
  public updateStepStatus(commandId: string, stepIndex: number, status: 'pending' | 'running' | 'completed' | 'failed', error?: string): void {
    const command = this.commands.get(commandId);
    if (command && command.steps[stepIndex]) {
      const step = command.steps[stepIndex];
      step.status = status;
      
      if (error) {
        step.error = error;
      }
      
      // Update command progress
      this.updateCommandProgress(command);
      
      this.notifyListeners(command);
    }
  }

  /**
   * Update an operation's status
   * @param commandId The command ID
   * @param stepIndex The step index
   * @param operationIndex The operation index
   * @param status The new status
   * @param error Optional error message
   */
  public updateOperationStatus(
    commandId: string, 
    stepIndex: number, 
    operationIndex: number, 
    status: 'pending' | 'running' | 'completed' | 'failed', 
    error?: string
  ): void {
    const command = this.commands.get(commandId);
    if (command && command.steps[stepIndex] && command.steps[stepIndex].operations[operationIndex]) {
      const operation = command.steps[stepIndex].operations[operationIndex];
      operation.status = status;
      
      if (error) {
        operation.error = error;
      }
      
      // Check if all operations in the step are completed
      this.checkStepCompletion(command, stepIndex);
      
      // Update command progress
      this.updateCommandProgress(command);
      
      this.notifyListeners(command);
    }
  }

  /**
   * Execute a command
   * @param commandId The command ID
   */
  /**
   * Set the workbook state manager for cache invalidation
   * @param workbookStateManager The workbook state manager
   */
  public setWorkbookStateManager(workbookStateManager: ClientWorkbookStateManager): void {
    this.workbookStateManager = workbookStateManager;
  }

  /**
   * Invalidate the workbook state cache if available
   * This should be called whenever the workbook is modified
   * @param operationType Optional operation type to selectively invalidate based on operation
   */
  private invalidateWorkbookCache(operationType?: string): void {
    if (this.workbookStateManager) {
      console.log('%c Checking cache invalidation for operation type: ' + (operationType || 'unknown'), 'color: #3498db');
      this.workbookStateManager.invalidateCache(operationType);
    }
  }

  public async executeCommand(commandId: string): Promise<void> {
    console.log(`🔍 [ClientCommandManager] executeCommand called with ID: ${commandId}`);
    
    const command = this.getCommand(commandId);
    if (!command) {
      throw new Error(`Command with ID ${commandId} not found`);
    }
    
    console.log(`📋 [ClientCommandManager] Executing command: ${command.description} (ID: ${commandId})`);
    console.log(`🔢 [ClientCommandManager] Command has ${command.steps.length} steps with ${command.steps.reduce((total, step) => total + step.operations.length, 0)} total operations`);
    
    try {
      // Update command status to running
      this.updateCommandStatus(commandId, CommandStatus.Running);
      
      // Check if this is an Excel DSL command by looking for operations with 'op' property in their value
      const hasExcelOperations = command.steps.some(step => 
        step.operations.some(op => 
          op.value && typeof op.value === 'object' && 'op' in op.value
        )
      );
      
      // Get operation types for selective cache invalidation
      let operationTypes: string[] = [];
      if (hasExcelOperations) {
        // Extract all operation types from the command
        command.steps.forEach(step => {
          step.operations.forEach(operation => {
            if (operation.value && typeof operation.value === 'object' && 'op' in operation.value) {
              const opType = operation.value.op as string;
              if (opType && !operationTypes.includes(opType)) {
                operationTypes.push(opType);
              }
            }
          });
        });
      }
      
      // Log the operation types for debugging
      if (operationTypes.length > 0) {
        console.log(`Command contains operation types: ${operationTypes.join(', ')}`);
      }
      
      if (hasExcelOperations) {
        // Use the Excel command adapter for DSL operations
        try {
          // Execute the command and get the actual operation types that were executed
          const executedOperationTypes = await this.excelCommandAdapter.executeCommand(command);
          
          // Mark all steps and operations as completed
          for (let i = 0; i < command.steps.length; i++) {
            this.updateStepStatus(commandId, i, 'completed');
            
            const step = command.steps[i];
            for (let j = 0; j < step.operations.length; j++) {
              this.updateOperationStatus(commandId, i, j, 'completed');
            }
          }
          
          // Log the actually executed operation types
          console.log(`Actually executed operation types: ${executedOperationTypes.join(', ')}`);
          
          // Invalidate cache based on operation types that were actually executed
          if (executedOperationTypes.length > 0) {
            // Pass all executed operation types as a comma-separated list for selective invalidation
            this.invalidateWorkbookCache(executedOperationTypes.join(','));
          } else {
            // If no operation types were executed, use the detected types as a fallback
            if (operationTypes.length > 0) {
              this.invalidateWorkbookCache(operationTypes.join(','));
            } else {
              // If no operation types found at all, invalidate normally
              this.invalidateWorkbookCache();
            }
          }
          
          // Update command status to completed
          this.updateCommandStatus(commandId, CommandStatus.Completed);
          
          // Invalidate cache after command execution completes
          this.invalidateWorkbookCache();
        } catch (error) {
          // Update command status to failed
          this.updateCommandStatus(commandId, CommandStatus.Failed, error.message);
          console.error(`Error executing Excel operations: ${error.message}`);
        }
      } else {
        // Use the traditional command executor for legacy operations
        // Execute each step in sequence
        for (let i = 0; i < command.steps.length; i++) {
          const step = command.steps[i];
          
          try {
            // Update step status to running
            this.updateStepStatus(commandId, i, 'running');
            
            // Execute each operation in the step
            for (let j = 0; j < step.operations.length; j++) {
              const operation = step.operations[j];
              
              try {
                // Update operation status to running
                this.updateOperationStatus(commandId, i, j, 'running');
                
                // Execute the operation
                await this.commandExecutor.executeOperation(operation);
                
                // Update operation status to completed
                this.updateOperationStatus(commandId, i, j, 'completed');
              } catch (error) {
                // Update operation status to failed
                this.updateOperationStatus(commandId, i, j, 'failed', error.message);
                
                // Don't throw here, try to continue with other operations
                console.error(`Error executing operation: ${error.message}`);
              }
            }
            
            // Check if all operations completed successfully
            if (step.operations.every(op => op.status === 'completed')) {
              this.updateStepStatus(commandId, i, 'completed');
            } else if (step.operations.some(op => op.status === 'failed')) {
              this.updateStepStatus(commandId, i, 'failed', 'One or more operations failed');
            }
          } catch (error) {
            // Update step status to failed
            this.updateStepStatus(commandId, i, 'failed', error.message);
            
            // Don't throw here, try to continue with other steps
            console.error(`Error executing step: ${error.message}`);
          }
        }
        
        // Check if all steps completed successfully
        if (command.steps.every(step => step.status === 'completed')) {
          this.updateCommandStatus(commandId, CommandStatus.Completed);
          // Invalidate cache after command execution completes successfully
          this.invalidateWorkbookCache();
        } else if (command.steps.some(step => step.status === 'failed')) {
          this.updateCommandStatus(commandId, CommandStatus.Failed, 'One or more steps failed');
          // Invalidate cache even on failure as the workbook might have been partially modified
          this.invalidateWorkbookCache();
        }
      }
    } catch (error) {
      // Update command status to failed
      this.updateCommandStatus(commandId, CommandStatus.Failed, error.message);
      console.error(`Error executing command: ${error.message}`);
    }
  }

  /**
   * Register a listener for command updates
   * @param listener The listener function
   * @returns A function to unregister the listener
   */
  public onCommandUpdate(listener: (command: Command) => void): () => void {
    this.commandUpdateListeners.push(listener);
    return () => {
      this.commandUpdateListeners = this.commandUpdateListeners.filter(l => l !== listener);
    };
  }

  /**
   * Check if all operations in a step are completed
   * @param command The command
   * @param stepIndex The step index
   */
  private checkStepCompletion(command: Command, stepIndex: number): void {
    const step = command.steps[stepIndex];
    const operations = step.operations;
    
    if (operations.every(op => op.status === 'completed')) {
      step.status = 'completed';
    } else if (operations.some(op => op.status === 'failed')) {
      step.status = 'failed';
      step.error = 'One or more operations failed';
    }
  }

  /**
   * Update a command's progress based on completed steps
   * @param command The command to update
   */
  private updateCommandProgress(command: Command): void {
    const totalSteps = command.steps.length;
    if (totalSteps === 0) {
      command.progress = 0;
      return;
    }
    
    let completedSteps = 0;
    let partialSteps = 0;
    
    for (const step of command.steps) {
      if (step.status === 'completed') {
        completedSteps++;
      } else if (step.status === 'running') {
        // Calculate partial completion for running steps
        const totalOps = step.operations.length;
        if (totalOps > 0) {
          const completedOps = step.operations.filter(op => op.status === 'completed').length;
          partialSteps += completedOps / totalOps;
        }
      }
    }
    
    // Calculate progress as percentage
    command.progress = Math.round(((completedSteps + partialSteps) / totalSteps) * 100);
  }

  /**
   * Notify all listeners of a command update
   * @param command The updated command
   */
  private notifyListeners(command: Command): void {
    for (const listener of this.commandUpdateListeners) {
      try {
        listener(command);
      } catch (error) {
        console.error('Error in command update listener:', error);
      }
    }
  }
}
