/**
 * Version History Provider
 * 
 * Provides a centralized service for initializing and managing version history components.
 * Acts as a facade for the version history system.
 */

import { v4 as uuidv4 } from 'uuid';
import { 
  WorkbookVersion, 
  WorkbookAction, 
  VersionEventType,
  CreateVersionOptions,
  RestoreVersionOptions,
  RestoreResult
} from '../../models/VersionModels';
import { VersionHistoryService } from './VersionHistoryService';
import { ActionRecorder } from './ActionRecorder';
import { VersionRestorer } from './VersionRestorer';
import { ClientExcelCommandInterpreter } from '../ClientExcelCommandInterpreter';
import { multimodalAnalysisService } from '../document-understanding/MultimodalAnalysisService';

/**
 * Provider for version history services
 */
export class VersionHistoryProvider {
  private versionHistoryService: VersionHistoryService;
  private actionRecorder: ActionRecorder;
  private versionRestorer: VersionRestorer;
  private commandInterpreter: ClientExcelCommandInterpreter | null = null;
  private currentWorkbookId: string = '';
  
  // Tracking for formatting changes
  private lastCheckedActionTimestamp: number = 0;
  private formatChangeCheckInterval: number | null = null;
  private lastFormattingActionIds: Set<string> = new Set();
  
  constructor() {
    // Initialize the version history components using the singleton instance
    this.versionHistoryService = VersionHistoryService.getInstance();
    this.versionHistoryService.initialize(); // Explicitly initialize the service
    this.actionRecorder = new ActionRecorder(this.versionHistoryService);
    this.versionRestorer = new VersionRestorer(this.versionHistoryService);
  }
  
  /**
   * Initialize the version history system with the command interpreter
   * @param commandInterpreter The Excel command interpreter
   */
  public initialize(commandInterpreter: ClientExcelCommandInterpreter): void {
    this.commandInterpreter = commandInterpreter;
    
    // Set the action recorder in the command interpreter
    commandInterpreter.setActionRecorder(this.actionRecorder);
    
    // Set up Office.js event handlers to detect sheet changes
    this.setupOfficeEventHandlers();
    
    console.log('Version history system initialized with Office.js change detection');
  }
  
  /**
   * Sets up Office.js event handlers to detect workbook changes
   */
  private setupOfficeEventHandlers(): void {
    // Ensure Office is initialized
    if (!Office || !Office.context || !Office.context.document) {
      console.warn('Office.js not initialized, cannot set up event handlers');
      return;
    }
    
    try {
      // Listen for selection changes as an indicator of user activity
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        this.handleSelectionChange.bind(this)
      );
      
      // When Excel is ready, set up additional event handlers
      Excel.run(async (context) => {
        // Get the active worksheet and workbook
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load('name');
        
        // Set up event handlers for worksheet activation
        context.workbook.worksheets.onActivated.add(this.handleWorksheetActivation.bind(this));
        
        // Set up event handlers for data changed events
        context.workbook.worksheets.onChanged.add(this.handleWorksheetDataChanged.bind(this));
        
        // Set up event handlers for formatting changes
        // Note: There's no direct event for formatting changes, so we'll use the selection changed
        // event as a proxy and check if formatting has changed when the selection changes
        
        await context.sync();
        console.log(`Set up Office.js event handlers for worksheet: ${worksheet.name}`);
      }).catch(error => {
        console.error('Error setting up Office.js event handlers:', error);
      });
    } catch (error) {
      console.error('Error setting up Office.js event handlers:', error);
    }
  }
  
  /**
   * Handles selection change events in the document
   * @param _eventArgs The event arguments (unused but required by the event handler signature)
   */
  private handleSelectionChange(_eventArgs: Office.DocumentSelectionChangedEventArgs): void {
    // Use selection changes as a potential indicator of formatting changes
    // We'll periodically check if we need to refresh images after multiple selection changes
    
    // Throttle the checks to avoid excessive processing
    if (this.formatChangeCheckInterval) {
      clearTimeout(this.formatChangeCheckInterval);
    }
    
    // Set a timeout to check for formatting changes after a brief delay
    this.formatChangeCheckInterval = window.setTimeout(() => {
      this.checkForFormattingChanges();
    }, 2000); // 2 second delay
  }
  
  /**
   * Handles worksheet activation events
   * @param event The worksheet activation event
   */
  private async handleWorksheetActivation(event: Excel.WorksheetActivatedEventArgs): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(event.worksheetId);
        worksheet.load('name');
        await context.sync();
        
        console.log(`Worksheet activated: ${worksheet.name}`);
        
        // When a worksheet is activated, check if we need to refresh its images
        if (this.currentWorkbookId) {
          // Refresh the images for this sheet if it's not already in progress
          await multimodalAnalysisService.refreshSheetImages(this.currentWorkbookId, worksheet.name);
        }
      });
    } catch (error) {
      console.error('Error handling worksheet activation:', error);
    }
  }
  
  /**
   * Handles worksheet data changed events
   * @param event The worksheet changed event
   */
  private async handleWorksheetDataChanged(event: Excel.WorksheetChangedEventArgs): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(event.worksheetId);
        worksheet.load('name');
        await context.sync();
        
        console.log(`Data changed in worksheet: ${worksheet.name}`);
        
        // When data changes in a worksheet, check if we need to refresh its images
        if (this.currentWorkbookId) {
          // Refresh the images for this sheet
          await multimodalAnalysisService.refreshSheetImages(this.currentWorkbookId, worksheet.name);
        }
      });
    } catch (error) {
      console.error('Error handling worksheet data change:', error);
    }
  }
  
  /**
   * Checks for formatting changes in the active worksheet
   */
  private async checkForFormattingChanges(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load('name');
        await context.sync();
        
        // If we have a current workbook ID, refresh the images for this sheet
        if (this.currentWorkbookId) {
          await multimodalAnalysisService.refreshSheetImages(this.currentWorkbookId, worksheet.name);
        }
      });
    } catch (error) {
      console.error('Error checking for formatting changes:', error);
    }
  }
  
  /**
   * Determines if an action is related to formatting changes
   * @param action The workbook action to check
   * @returns True if the action is related to formatting
   */
  private isFormattingRelatedAction(action: WorkbookAction): boolean {
    // Check if the action type is related to formatting
    if (action.type === VersionEventType.FormatOperation) {
      return true;
    }
    
    // Check if the operation contains formatting-related properties
    const formattingOperations = [
      'format_range',
      'conditional_format',
      'merge_cells',
      'unmerge_cells',
      'set_font',
      'set_fill',
      'set_border',
      'set_number_format',
      'set_alignment',
      'create_table',
      'format_table',
      'format_chart'
    ];
    
    // Check if the operation type is in the list of formatting operations
    if (action.operation && action.operation.op && formattingOperations.includes(action.operation.op)) {
      return true;
    }
    
    // Check if the beforeState contains formatting information
    if (action.beforeState && action.beforeState.formats) {
      return true;
    }
    
    // Check if the operation metadata contains formatting-related properties
    if (action.metadata) {
      const formattingProperties = [
        'format',
        'style',
        'fill',
        'font',
        'border',
        'numberFormat',
        'alignment',
        'color',
        'background',
        'bold',
        'italic',
        'underline'
      ];
      
      // Check if any formatting properties are present in the metadata
      return formattingProperties.some(prop => 
        action.metadata && typeof action.metadata === 'object' && prop in action.metadata
      );
    }
    
    return false;
  }
  
  /**
   * Set the current workbook ID
   * @param workbookId The ID of the current workbook
   */
  public setCurrentWorkbookId(workbookId: string): void {
    console.log(`Setting current workbook ID in VersionHistoryProvider: ${workbookId}`);
    this.currentWorkbookId = workbookId;
    
    // Also set it in the command interpreter if available
    if (this.commandInterpreter) {
      this.commandInterpreter.setCurrentWorkbookId(workbookId);
    }
    
    // When the workbook ID changes, set up the Office.js event handlers again
    // to ensure they're tracking the correct workbook
    this.setupOfficeEventHandlers();
  }
  
  /**
   * Create a new workbook ID if one doesn't exist
   * @returns The workbook ID
   */
  public ensureWorkbookId(): string {
    if (!this.currentWorkbookId) {
      this.currentWorkbookId = uuidv4();
      
      // Also set it in the command interpreter if available
      if (this.commandInterpreter) {
        this.commandInterpreter.setCurrentWorkbookId(this.currentWorkbookId);
      }
      
      console.log(`Generated new workbook ID: ${this.currentWorkbookId}`);
    }
    
    return this.currentWorkbookId;
  }
  
  /**
   * Create a new version point in the history
   * @param options Options for creating the version
   * @returns The ID of the created version
   */
  public createVersion(options: CreateVersionOptions = {}): string {
    const workbookId = this.ensureWorkbookId();
    return this.versionHistoryService.createVersion(workbookId, options);
  }
  
  // Removed duplicate getAllActions method - using the one below
  
  /**
   * Get all versions for the current workbook
   * @returns Array of workbook versions
   */
  public getVersions(): WorkbookVersion[] {
    if (!this.currentWorkbookId) {
      return [];
    }
    
    return this.versionHistoryService.getVersionsForWorkbook(this.currentWorkbookId);
  }
  
  /**
   * Get a specific version by ID
   * @param versionId The ID of the version
   * @returns The workbook version or undefined if not found
   */
  public getVersion(versionId: string): WorkbookVersion | undefined {
    return this.versionHistoryService.getVersion(versionId);
  }
  
  /**
   * Get all actions for a specific version
   * @param versionId ID of the version to get actions for
   * @returns Array of actions associated with the version
   */
  getActionsForVersion(versionId: string): WorkbookAction[] {
    const version = this.versionHistoryService.getVersion(versionId);
    if (!version) return [];
    
    const allActions = this.versionHistoryService.getAllActions();
    return version.actionIds
      .map(id => allActions.get(id))
      .filter(action => action !== undefined) as WorkbookAction[];
  }
  
  /**
   * Get all actions for a specific action group
   * @param groupId ID of the action group to get actions for
   * @returns Array of actions associated with the action group
   */
  getActionsForActionGroup(groupId: string): WorkbookAction[] {
    // Extract timestamp from the action group ID
    if (!groupId.startsWith('action-group-')) return [];
    
    const timestamp = parseInt(groupId.replace('action-group-', ''));
    if (isNaN(timestamp)) return [];
    
    // Find actions within a small time window of the timestamp
    const allActions = Array.from(this.versionHistoryService.getAllActions().values());
    const GROUPING_THRESHOLD_MS = 1000; // 1 second window
    
    return allActions.filter(action => 
      Math.abs(action.timestamp - timestamp) < GROUPING_THRESHOLD_MS &&
      // Also check if they have the same query ID in metadata (if available)
      (!action.metadata?.queryId || 
       allActions.some(a => 
         Math.abs(a.timestamp - timestamp) < 10 && // Very close to the timestamp
         a.metadata?.queryId === action.metadata?.queryId
       ))
    );
  }
  
  /**
   * Get all actions for the current workbook
   * @returns Array of all workbook actions
   */
  public getAllActions(): WorkbookAction[] {
    if (!this.currentWorkbookId) {
      console.warn('Cannot get actions: No current workbook ID');
      return [];
    }
    
    const actions = this.versionHistoryService.getActionsForWorkbook(this.currentWorkbookId);
    console.log(`💾 [VersionHistoryProvider] Retrieved ${actions.length} actions for workbook: ${this.currentWorkbookId}`);
    return actions;
  }
  
  /**
   * Get the version history service instance
   * @returns The version history service instance
   */
  public getVersionHistoryService(): VersionHistoryService {
    return this.versionHistoryService;
  }
  

  
  /**
   * Restore a workbook to a previous version
   * @param options Options for the restore operation
   * @returns Result of the restore operation
   */
  public async restoreVersion(options: RestoreVersionOptions): Promise<RestoreResult> {
    return this.versionRestorer.restoreVersion(options);
  }
  
  /**
   * Clear all version history for the current workbook
   * @returns True if successful, false otherwise
   */
  public clearVersionHistory(): boolean {
    if (!this.currentWorkbookId) {
      console.error('Cannot clear version history: No current workbook ID');
      return false;
    }
    
    try {
      this.versionHistoryService.clearVersionHistory(this.currentWorkbookId);
      return true;
    } catch (error) {
      console.error('Error clearing version history:', error);
      return false;
    }
  }
}
