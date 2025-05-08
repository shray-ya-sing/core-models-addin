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

/**
 * Provider for version history services
 */
export class VersionHistoryProvider {
  private versionHistoryService: VersionHistoryService;
  private actionRecorder: ActionRecorder;
  private versionRestorer: VersionRestorer;
  private commandInterpreter: ClientExcelCommandInterpreter | null = null;
  private currentWorkbookId: string = '';
  
  constructor() {
    // Initialize the version history components
    this.versionHistoryService = new VersionHistoryService();
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
    
    console.log('Version history system initialized');
  }
  
  /**
   * Set the current workbook ID
   * @param workbookId The ID of the current workbook
   */
  public setCurrentWorkbookId(workbookId: string): void {
    this.currentWorkbookId = workbookId;
    
    // Also set it in the command interpreter if available
    if (this.commandInterpreter) {
      this.commandInterpreter.setCurrentWorkbookId(workbookId);
    }
    
    console.log(`Current workbook ID set to: ${workbookId}`);
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
   * Search for versions matching a query
   * @param query The search query
   * @returns Array of matching versions
   */
  public searchVersions(query: string): WorkbookVersion[] {
    if (!this.currentWorkbookId) {
      return [];
    }
    
    return this.versionHistoryService.searchVersions(this.currentWorkbookId, query);
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
