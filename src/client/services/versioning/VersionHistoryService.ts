/**
 * Version History Service
 * 
 * Implements action-based version tracking for Excel workbooks.
 * Records operations as they happen and provides functionality to restore previous versions.
 */

import { v4 as uuidv4 } from 'uuid';
import { ExcelOperation, ExcelOperationType } from '../../models/ExcelOperationModels';
import { 
  WorkbookAction, 
  WorkbookVersion, 
  VersionEventType,
  AffectedRange,
  BeforeState,
  CreateVersionOptions,
  RestoreVersionOptions,
  RestoreResult
} from '../../models/VersionModels';

/**
 * Service for managing version history of Excel workbooks
 * Implemented as a singleton to ensure only one instance exists throughout the application
 */
export class VersionHistoryService {
  // Singleton instance
  private static instance: VersionHistoryService;
  
  private actions: Map<string, WorkbookAction> = new Map();
  private versions: Map<string, WorkbookVersion> = new Map();
  private workbookVersions: Map<string, string[]> = new Map(); // workbookId -> versionIds
  private workbookActions: Map<string, string[]> = new Map();  // workbookId -> actionIds
  
  private storageKeyPrefix = 'excel-addin-version-';
  private actionsKey = 'actions';
  private versionsKey = 'versions';
  private workbookVersionsKey = 'workbook-versions';
  private workbookActionsKey = 'workbook-actions';
  private initialized = false;
  
  /**
   * Private constructor to prevent direct instantiation
   */
  private constructor() {
    // Initialization will be done in initialize() method
  }
  
  /**
   * Get the singleton instance of VersionHistoryService
   * @returns The singleton instance
   */
  public static getInstance(): VersionHistoryService {
    if (!VersionHistoryService.instance) {
      VersionHistoryService.instance = new VersionHistoryService();
    }
    return VersionHistoryService.instance;
  }
  
  /**
   * Initialize the service by loading data from storage
   * This is separated from the constructor to allow controlled initialization
   */
  public initialize(): void {
    if (!this.initialized) {
      this.loadFromStorage();
      this.initialized = true;
    }
  }
  
  /**
   * Load version history data from localStorage
   */
  private loadFromStorage(): void {
    try {
      const startTime = performance.now();
      console.log(`🔍 [VersionHistoryService] Loading version history from localStorage...`);
      
      // Load actions
      const actionsJson = localStorage.getItem(`${this.storageKeyPrefix}${this.actionsKey}`);
      if (actionsJson) {
        const actionsArray = JSON.parse(actionsJson) as WorkbookAction[];
        this.actions = new Map(actionsArray.map(action => [action.id, action]));
        console.log(`📚 [VersionHistoryService] Loaded ${actionsArray.length} actions from localStorage`);
        
        // Group actions by workbook for better visibility
        const actionsByWorkbook = actionsArray.reduce((acc, action) => {
          if (!acc[action.workbookId]) {
            acc[action.workbookId] = [];
          }
          acc[action.workbookId].push(action.id);
          return acc;
        }, {} as Record<string, string[]>);
        
        console.log(`📁 [VersionHistoryService] Actions by workbook:`, 
          Object.entries(actionsByWorkbook).map(([wbId, actions]) => 
            `${wbId}: ${actions.length} actions`
          )
        );
      } else {
        console.log(`⚠️ [VersionHistoryService] No actions found in localStorage`);
        this.actions = new Map();
      }
      
      // Load versions
      const versionsJson = localStorage.getItem(`${this.storageKeyPrefix}${this.versionsKey}`);
      if (versionsJson) {
        const versionsArray = JSON.parse(versionsJson) as WorkbookVersion[];
        this.versions = new Map(versionsArray.map(version => [version.id, version]));
        console.log(`📜 [VersionHistoryService] Loaded ${versionsArray.length} versions from localStorage`);
        
        // Group versions by workbook for better visibility
        const versionsByWorkbook = versionsArray.reduce((acc, version) => {
          if (!acc[version.workbookId]) {
            acc[version.workbookId] = [];
          }
          acc[version.workbookId].push(version.id);
          return acc;
        }, {} as Record<string, string[]>);
        
        console.log(`📂 [VersionHistoryService] Versions by workbook:`, 
          Object.entries(versionsByWorkbook).map(([wbId, versions]) => 
            `${wbId}: ${versions.length} versions`
          )
        );
      } else {
        console.log(`⚠️ [VersionHistoryService] No versions found in localStorage`);
        this.versions = new Map();
      }
      
      // Load workbook versions mapping
      const workbookVersionsJson = localStorage.getItem(`${this.storageKeyPrefix}${this.workbookVersionsKey}`);
      if (workbookVersionsJson) {
        const workbookVersionsObj = JSON.parse(workbookVersionsJson);
        this.workbookVersions = new Map(Object.entries(workbookVersionsObj));
        console.log(`📃 [VersionHistoryService] Loaded version mappings for ${Object.keys(workbookVersionsObj).length} workbooks`);
      } else {
        console.log(`⚠️ [VersionHistoryService] No workbook version mappings found in localStorage`);
        this.workbookVersions = new Map();
      }
      
      // Load workbook actions mapping
      const workbookActionsJson = localStorage.getItem(`${this.storageKeyPrefix}${this.workbookActionsKey}`);
      if (workbookActionsJson) {
        const workbookActionsObj = JSON.parse(workbookActionsJson);
        this.workbookActions = new Map(Object.entries(workbookActionsObj));
        console.log(`📄 [VersionHistoryService] Loaded action mappings for ${Object.keys(workbookActionsObj).length} workbooks`);
        
        // Log action counts per workbook
        Object.entries(workbookActionsObj).forEach(([workbookId, actions]) => {
          console.log(`📊 [VersionHistoryService] Workbook ${workbookId} has ${(actions as string[]).length} actions`);
        });
      } else {
        console.log(`⚠️ [VersionHistoryService] No workbook action mappings found in localStorage`);
        this.workbookActions = new Map();
      }
      
      const endTime = performance.now();
      console.log(`✅ [VersionHistoryService] Loaded version history: ${this.actions.size} actions, ${this.versions.size} versions in ${(endTime - startTime).toFixed(2)}ms`);
    } catch (error) {
      console.error('❌ [VersionHistoryService] Error loading version history from storage:', error);
      // Initialize empty maps if loading fails
      this.actions = new Map();
      this.versions = new Map();
      this.workbookVersions = new Map();
      this.workbookActions = new Map();
      console.log(`⚠️ [VersionHistoryService] Initialized empty maps due to loading error`);
    }
  }
  
  /**
   * Save version history data to localStorage
   */
  private saveToStorage(): void {
    try {
      const startTime = performance.now();
      console.log(`💾 [VersionHistoryService] Saving data to localStorage...`);
      
      // Save actions
      const actionsArray = Array.from(this.actions.values());
      console.log(`📁 [VersionHistoryService] Saving ${actionsArray.length} actions to localStorage`);
      localStorage.setItem(`${this.storageKeyPrefix}${this.actionsKey}`, JSON.stringify(actionsArray));
      
      // Save versions
      const versionsArray = Array.from(this.versions.values());
      console.log(`📂 [VersionHistoryService] Saving ${versionsArray.length} versions to localStorage`);
      localStorage.setItem(`${this.storageKeyPrefix}${this.versionsKey}`, JSON.stringify(versionsArray));
      
      // Save workbook versions mapping
      const workbookVersionsObj = Object.fromEntries(this.workbookVersions);
      const workbookCount = Object.keys(workbookVersionsObj).length;
      console.log(`📃 [VersionHistoryService] Saving version mappings for ${workbookCount} workbooks`);
      localStorage.setItem(`${this.storageKeyPrefix}${this.workbookVersionsKey}`, JSON.stringify(workbookVersionsObj));
      
      // Save workbook actions mapping
      const workbookActionsObj = Object.fromEntries(this.workbookActions);
      console.log(`📄 [VersionHistoryService] Saving action mappings for ${Object.keys(workbookActionsObj).length} workbooks`);
      localStorage.setItem(`${this.storageKeyPrefix}${this.workbookActionsKey}`, JSON.stringify(workbookActionsObj));
      
      const endTime = performance.now();
      console.log(`✅ [VersionHistoryService] Successfully saved all data to localStorage in ${(endTime - startTime).toFixed(2)}ms`);
    } catch (error) {
      console.error('❌ [VersionHistoryService] Error saving version history to storage:', error);
    }
  }
  
  /**
   * Record an Excel operation as an action in the version history
   * @param workbookId The ID of the workbook
   * @param operation The Excel operation being performed
   * @param beforeState The state before the operation was performed
   * @param affectedRanges The ranges affected by the operation
   * @returns The ID of the recorded action
   */
  public recordAction(
    workbookId: string, 
    operation: ExcelOperation, 
    beforeState: BeforeState, 
    affectedRanges: AffectedRange[]
  ): string {
    const startTime = performance.now();
    console.log(`💾 [VersionHistoryService] Recording action for workbook: ${workbookId}, operation: ${operation.op}`);
    
    try {
      // Generate a unique ID for this action
      const actionId = uuidv4();
      console.log(`🆔 [VersionHistoryService] Generated action ID: ${actionId}`);
      
      // Create a human-readable description based on the operation type
      const description = this.generateActionDescription(operation);
      console.log(`📝 [VersionHistoryService] Action description: ${description}`);
      
      // Determine the version event type based on the operation
      const type = this.determineEventType(operation);
      console.log(`🏷️ [VersionHistoryService] Action type: ${type}`);
      
      // Log before state details with more information
      const hasValues = beforeState.values && Array.isArray(beforeState.values) && beforeState.values.length > 0;
      const hasFormulas = beforeState.formulas && Array.isArray(beforeState.formulas) && beforeState.formulas.length > 0;
      const formatCount = beforeState.formats ? beforeState.formats.length : 0;
      const sheetPropertiesCount = Object.keys(beforeState.sheetProperties || {}).length;
      
      console.log(`📊 [VersionHistoryService] Before state details:`, {
        hasValues,
        valueRows: hasValues ? beforeState.values.length : 0,
        valueColumns: hasValues && beforeState.values.length > 0 ? beforeState.values[0].length : 0,
        hasFormulas,
        formulaRows: hasFormulas ? beforeState.formulas.length : 0,
        formulaColumns: hasFormulas && beforeState.formulas.length > 0 ? beforeState.formulas[0].length : 0,
        formatCount,
        sheetPropertiesCount,
        affectedRangesCount: affectedRanges.length,
        affectedRanges: affectedRanges.map(r => `${r.sheetName}!${r.range || '[sheet-level]'}`)
      });
      
      // Create the action object
      const action: WorkbookAction = {
        id: actionId,
        workbookId,
        timestamp: Date.now(),
        type,
        operation,
        description,
        affectedRanges,
        beforeState
      };
      
      // Store the action
      this.actions.set(actionId, action);
      console.log(`✅ [VersionHistoryService] Action stored in memory with ID: ${actionId}`);
      
      // Add to workbook actions mapping
      if (!this.workbookActions.has(workbookId)) {
        this.workbookActions.set(workbookId, []);
        console.log(`🌐 [VersionHistoryService] Created new action list for workbook: ${workbookId}`);
      }
      this.workbookActions.get(workbookId)?.push(actionId);
      console.log(`📎 [VersionHistoryService] Added action ${actionId} to workbook ${workbookId} action list`);
      
      // Get current action count for this workbook
      const actionCount = this.workbookActions.get(workbookId)?.length || 0;
      console.log(`📈 [VersionHistoryService] Workbook ${workbookId} now has ${actionCount} recorded actions`);
      
      // Save to storage immediately after recording the action
      console.log(`💾 [VersionHistoryService] Saving action to persistent storage...`);
      this.saveToStorage();
      
      const endTime = performance.now();
      console.log(`✅ [VersionHistoryService] Action recording completed in ${(endTime - startTime).toFixed(2)}ms`);
      
      return actionId;
    } catch (error) {
      const endTime = performance.now();
      console.error(`❌ [VersionHistoryService] Error recording action:`, {
        error: error.message,
        stack: error.stack,
        workbookId,
        operationType: operation.op,
        duration: `${(endTime - startTime).toFixed(2)}ms`,
        affectedRangesCount: affectedRanges.length
      });
      throw error; // Re-throw to allow caller to handle
    }
  }
  
  /**
   * Create a new version point in the history
   * @param workbookId The ID of the workbook
   * @param options Options for creating the version
   * @returns The ID of the created version
   */
  public createVersion(workbookId: string, options: CreateVersionOptions = {}): string {
    // Generate a unique ID for this version
    const versionId = uuidv4();
    
    // Set default values
    const type = options.type || VersionEventType.ManualSave;
    const description = options.description || this.generateVersionDescription(type);
    const author = options.author || 'User';
    const timestamp = Date.now();
    
    // Determine which actions to include in this version
    let actionIds: string[] = [];
    
    if (options.includeActionsSince) {
      // Include actions since the specified timestamp
      const workbookActionIds = this.workbookActions.get(workbookId) || [];
      actionIds = workbookActionIds.filter(id => {
        const action = this.actions.get(id);
        return action && action.timestamp >= options.includeActionsSince!;
      });
    } else {
      // Include all actions since the last version
      const workbookVersions = this.workbookVersions.get(workbookId) || [];
      
      if (workbookVersions.length > 0) {
        // Get the most recent version
        const latestVersionId = workbookVersions[0];
        const latestVersion = this.versions.get(latestVersionId);
        
        if (latestVersion) {
          // Include actions that occurred after the latest version
          const workbookActionIds = this.workbookActions.get(workbookId) || [];
          actionIds = workbookActionIds.filter(id => {
            const action = this.actions.get(id);
            return action && action.timestamp > latestVersion.timestamp;
          });
        }
      } else {
        // No previous versions, include all actions for this workbook
        actionIds = this.workbookActions.get(workbookId) || [];
      }
    }
    
    // Create the version object
    const version: WorkbookVersion = {
      id: versionId,
      workbookId,
      timestamp,
      type,
      description,
      author,
      actionIds,
      tags: options.tags || []
    };
    
    // Store the version
    this.versions.set(versionId, version);
    
    // Add to workbook versions mapping
    if (!this.workbookVersions.has(workbookId)) {
      this.workbookVersions.set(workbookId, []);
    }
    
    // Add to the beginning of the array (most recent first)
    const workbookVersionIds = this.workbookVersions.get(workbookId)!;
    workbookVersionIds.unshift(versionId);
    
    // Save to storage
    this.saveToStorage();
    
    console.log(`Created version: ${description} with ${actionIds.length} actions (${versionId})`);
    
    return versionId;
  }
  
  /**
   * Get all versions for a specific workbook
   * @param workbookId The ID of the workbook
   * @returns Array of workbook versions, sorted by timestamp (newest first)
   */
  public getVersionsForWorkbook(workbookId: string): WorkbookVersion[] {
    const versionIds = this.workbookVersions.get(workbookId) || [];
    return versionIds
      .map(id => this.versions.get(id))
      .filter((version): version is WorkbookVersion => version !== undefined)
      .sort((a, b) => b.timestamp - a.timestamp);
  }
  
  /**
   * Get all actions for a specific workbook
   * @param workbookId The ID of the workbook
   * @returns Array of workbook actions
   */
  public getActionsForWorkbook(workbookId: string): WorkbookAction[] {
    const actionIds = this.workbookActions.get(workbookId) || [];
    console.log(`🔍 [VersionHistoryService] Getting ${actionIds.length} actions for workbook ${workbookId}`);
    return actionIds
      .map(id => this.actions.get(id))
      .filter((action): action is WorkbookAction => action !== undefined)
      .sort((a, b) => b.timestamp - a.timestamp); // Newest first
  }
  
  /**
   * Get a specific action by ID
   * @param actionId The ID of the action to retrieve
   * @returns The action or undefined if not found
   */
  public getAction(actionId: string): WorkbookAction | undefined {
    return this.actions.get(actionId);
  }
  
  /**
   * Get all actions
   * @returns Map of all actions
   */
  public getAllActions(): Map<string, WorkbookAction> {
    return this.actions;
  }
  
  /**
   * Get a specific version by ID
   * @param versionId The ID of the version
   * @returns The workbook version or undefined if not found
   */
  public getVersion(versionId: string): WorkbookVersion | undefined {
    return this.versions.get(versionId);
  }
  
  /**
   * Get all actions for a specific version
   * @param versionId The ID of the version
   * @returns Array of workbook actions in this version
   */
  public getActionsForVersion(versionId: string): WorkbookAction[] {
    const version = this.versions.get(versionId);
    if (!version) {
      return [];
    }
    
    return version.actionIds
      .map(id => this.actions.get(id))
      .filter((action): action is WorkbookAction => action !== undefined)
      .sort((a, b) => a.timestamp - b.timestamp); // Oldest first for restoration order
  }
  
  /**
   * Search for versions matching a query
   * @param workbookId The ID of the workbook
   * @param query The search query
   * @returns Array of matching versions
   */
  public searchVersions(workbookId: string, query: string): WorkbookVersion[] {
    if (!query) {
      return this.getVersionsForWorkbook(workbookId);
    }
    
    const normalizedQuery = query.toLowerCase();
    
    return this.getVersionsForWorkbook(workbookId).filter(version => {
      // Search in description
      if (version.description.toLowerCase().includes(normalizedQuery)) {
        return true;
      }
      
      // Search in author
      if (version.author.toLowerCase().includes(normalizedQuery)) {
        return true;
      }
      
      // Search in tags
      if (version.tags.some(tag => tag.toLowerCase().includes(normalizedQuery))) {
        return true;
      }
      
      // Search in actions
      const actions = this.getActionsForVersion(version.id);
      return actions.some(action => 
        action.description.toLowerCase().includes(normalizedQuery)
      );
    });
  }
  
  /**
   * Generate a human-readable description for an action based on the operation type
   * @param operation The Excel operation
   * @returns A human-readable description
   */
  private generateActionDescription(operation: ExcelOperation): string {
    const opType = operation.op as string;
    
    switch (opType) {
      case ExcelOperationType.SET_VALUE:
        return 'target' in operation ? `Set value in ${operation.target}` : 'Set value';
      
      case ExcelOperationType.ADD_FORMULA:
        return 'target' in operation ? `Add formula to ${operation.target}` : 'Add formula';
      
      case ExcelOperationType.FORMAT_RANGE:
        return 'range' in operation ? `Format range ${operation.range}` : 'Format range';
      
      case ExcelOperationType.CLEAR_RANGE:
        return 'range' in operation ? `Clear range ${operation.range}` : 'Clear range';
      
      case ExcelOperationType.CREATE_TABLE:
        return 'range' in operation ? `Create table in range ${operation.range}` : 'Create table';
      
      case ExcelOperationType.SORT_RANGE:
        return 'range' in operation ? `Sort range ${operation.range}` : 'Sort range';
      
      case ExcelOperationType.FILTER_RANGE:
        return 'range' in operation ? `Filter range ${operation.range}` : 'Filter range';
      
      case ExcelOperationType.CREATE_SHEET:
        return 'name' in operation ? `Create sheet "${operation.name}"` : 'Create sheet';
      
      case ExcelOperationType.DELETE_SHEET:
        return 'name' in operation ? `Delete sheet "${operation.name}"` : 'Delete sheet';
      
      case ExcelOperationType.RENAME_SHEET:
        return 'name' in operation ? `Rename sheet to "${operation.name}"` : 'Rename sheet';
      
      case ExcelOperationType.COPY_RANGE:
        return ('source' in operation && 'destination' in operation) ? 
          `Copy range from ${operation.source} to ${operation.destination}` : 'Copy range';
      
      case ExcelOperationType.MERGE_CELLS:
        return 'range' in operation ? `Merge cells in range ${operation.range}` : 'Merge cells';
      
      case ExcelOperationType.UNMERGE_CELLS:
        return 'range' in operation ? `Unmerge cells in range ${operation.range}` : 'Unmerge cells';
      
      case ExcelOperationType.CREATE_CHART:
        return ('type' in operation && 'range' in operation) ? 
          `Create ${operation.type} chart for range ${operation.range}` : 'Create chart';
      
      default:
        return `Performed ${operation.op} operation`;
    }
  }
  
  /**
   * Generate a human-readable description for a version based on its type
   * @param type The version event type
   * @returns A human-readable description
   */
  private generateVersionDescription(type: VersionEventType): string {
    switch (type) {
      case VersionEventType.ManualSave:
        return `Manual save at ${new Date().toLocaleString()}`;
      
      case VersionEventType.AutoSave:
        return `Auto save at ${new Date().toLocaleString()}`;
      
      case VersionEventType.Restore:
        return `Restored version at ${new Date().toLocaleString()}`;
      
      case VersionEventType.InitialState:
        return `Initial state at ${new Date().toLocaleString()}`;
      
      default:
        return `Version created at ${new Date().toLocaleString()}`;
    }
  }
  
  /**
   * Clear all version history for a specific workbook
   * @param workbookId The ID of the workbook to clear history for
   */
  public clearVersionHistory(workbookId: string): void {
    try {
      console.log(`Clearing version history for workbook: ${workbookId}`);
      
      // Get all versions for this workbook
      const versionIds = this.workbookVersions.get(workbookId) || [];
      
      // Get all actions for this workbook
      const actionIds = this.workbookActions.get(workbookId) || [];
      
      // Remove all versions
      versionIds.forEach(id => {
        this.versions.delete(id);
      });
      
      // Remove all actions
      actionIds.forEach(id => {
        this.actions.delete(id);
      });
      
      // Clear the workbook mappings
      this.workbookVersions.delete(workbookId);
      this.workbookActions.delete(workbookId);
      
      // Save changes to storage
      this.saveToStorage();
      
      console.log(`Successfully cleared version history for workbook: ${workbookId}`);
    } catch (error) {
      console.error('Error clearing version history:', error);
      throw error;
    }
  }
  
  /**
   * Determine the version event type based on the operation
   * @param operation The Excel operation
   * @returns The version event type
   */
  private determineEventType(operation: ExcelOperation): VersionEventType {
    const opType = operation.op as string;
    
    // Cell operations
    if (opType === ExcelOperationType.SET_VALUE || opType === ExcelOperationType.ADD_FORMULA) {
      return VersionEventType.CellOperation;
    }
    
    // Range operations
    if ([
      ExcelOperationType.FORMAT_RANGE,
      ExcelOperationType.CLEAR_RANGE,
      ExcelOperationType.SORT_RANGE,
      ExcelOperationType.FILTER_RANGE,
      ExcelOperationType.COPY_RANGE,
      ExcelOperationType.MERGE_CELLS,
      ExcelOperationType.UNMERGE_CELLS
    ].includes(opType as any)) {
      return VersionEventType.RangeOperation;
    }
    
    // Sheet operations
    if ([
      ExcelOperationType.CREATE_SHEET,
      ExcelOperationType.DELETE_SHEET,
      ExcelOperationType.RENAME_SHEET
    ].includes(opType as any)) {
      return VersionEventType.SheetOperation;
    }
    
    // Chart operations
    if ([
      ExcelOperationType.CREATE_CHART,
      ExcelOperationType.FORMAT_CHART
    ].includes(opType as any)) {
      return VersionEventType.ChartOperation;
    }
    
    // Table operations
    if (opType === ExcelOperationType.CREATE_TABLE) {
      return VersionEventType.TableOperation;
    }
    
    // Composite operations
    if ([
      ExcelOperationType.COMPOSITE_OPERATION,
      ExcelOperationType.BATCH_OPERATION
    ].includes(opType as any)) {
      return VersionEventType.CompositeOperation;
    }
    
    // Default
    return VersionEventType.WorkbookOperation;
  }
}
