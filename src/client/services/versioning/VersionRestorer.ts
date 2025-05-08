/**
 * Version Restorer
 * 
 * Provides functionality to restore previous versions of a workbook
 * by undoing actions in reverse chronological order.
 */

import { WorkbookAction, WorkbookVersion, RestoreVersionOptions, RestoreResult, VersionEventType } from '../../models/VersionModels';
import { VersionHistoryService } from './VersionHistoryService';
import { UndoHandlers } from './UndoHandlers';

// Ensure we have access to the Office.js API
declare const Excel: any;

/**
 * Service for restoring previous versions of a workbook
 */
export class VersionRestorer {
  private versionHistoryService: VersionHistoryService;
  private undoHandlers: UndoHandlers;
  
  constructor(versionHistoryService: VersionHistoryService) {
    this.versionHistoryService = versionHistoryService;
    this.undoHandlers = new UndoHandlers();
  }
  
  /**
   * Restore a workbook to a previous version
   * @param options Options for the restore operation
   * @returns Result of the restore operation
   */
  public async restoreVersion(options: RestoreVersionOptions): Promise<RestoreResult> {
    console.log(`🔄 [VersionRestorer] restoreVersion called with options:`, options);
    const { versionId, createRestorePoint = true, selective = false, selectiveRanges = [] } = options;
    
    // Get the version to restore
    const version = this.versionHistoryService.getVersion(versionId);
    if (!version) {
      console.error(`❌ [VersionRestorer] Version not found: ${versionId}`);
      return {
        success: false,
        restoredVersion: null,
        restoredActions: 0,
        errors: [`Version not found: ${versionId}`]
      };
    }
    
    console.log(`📋 [VersionRestorer] Found version to restore:`, version);
    
    // Create a restore point if requested
    let restorePointId: string | undefined;
    if (createRestorePoint) {
      console.log(`💾 [VersionRestorer] Creating restore point before reverting to version ${versionId}`);
      try {
        restorePointId = this.versionHistoryService.createVersion(version.workbookId, {
          description: `Restore point before reverting to ${version.description}`,
          type: VersionEventType.Restore,
          author: 'System'
        });
        console.log(`✅ [VersionRestorer] Restore point created with ID: ${restorePointId}`);
      } catch (error) {
        console.error(`❌ [VersionRestorer] Error creating restore point:`, error);
      }
    } else {
      console.log(`📝 [VersionRestorer] Skipping restore point creation as per options`);
    }
    
    // Get all actions for this version
    const actions = this.versionHistoryService.getActionsForVersion(versionId);
    console.log(`📋 [VersionRestorer] Retrieved ${actions.length} actions for version ${versionId}`);
    
    if (actions.length === 0) {
      console.warn(`⚠️ [VersionRestorer] No actions found for version ${versionId}`);
      return {
        success: false,
        restoredVersion: version,
        restoredActions: 0,
        errors: ['No actions found for this version'],
        restorePointId
      };
    }
    
    // Log action details for debugging
    console.log(`📋 [VersionRestorer] Actions to restore:`, 
      actions.map(a => ({
        id: a.id,
        type: a.operation?.op,
        description: a.description,
        ranges: a.affectedRanges?.map(r => `${r.sheetName}!${r.range}`)
      })));
    
    // If selective restore, filter actions to only those affecting the specified ranges
    let actionsToRestore = actions;
    if (selective && selectiveRanges.length > 0) {
      console.log(`🔍 [VersionRestorer] Performing selective restore for ranges:`, selectiveRanges);
      
      actionsToRestore = actions.filter(action => {
        return action.affectedRanges.some(affectedRange => {
          return selectiveRanges.some(selectiveRange => {
            return (
              affectedRange.sheetName === selectiveRange.sheetName &&
              affectedRange.range === selectiveRange.range
            );
          });
        });
      });
      
      console.log(`🔍 [VersionRestorer] Filtered to ${actionsToRestore.length} actions for selective restore`);
    } else {
      console.log(`📋 [VersionRestorer] Performing full restore with all ${actions.length} actions`);
    }
    
    // Restore the version by undoing actions in reverse chronological order
    const errors: string[] = [];
    let restoredCount = 0;
    
    try {
      // Check if Excel API is available
      if (typeof Excel === 'undefined') {
        console.error(`❌ [VersionRestorer] Excel API is not available. Make sure Office.js is properly loaded.`);
        throw new Error('Excel API is not available. Make sure Office.js is properly loaded.');
      }
      
      console.log(`🔄 [VersionRestorer] Starting version restore for version: ${versionId} with ${actionsToRestore.length} actions`);
      console.log(`🔄 [VersionRestorer] Excel API is available, proceeding with Excel.run()`);
      
      await Excel.run(async (context) => {
        console.log(`🔄 [VersionRestorer] Inside Excel.run context`);
        // Process actions in reverse order (newest to oldest)
        for (let i = actionsToRestore.length - 1; i >= 0; i--) {
          const action = actionsToRestore[i];
          
          try {
            console.log(`🔄 [VersionRestorer] Undoing action ${i+1}/${actionsToRestore.length}: ${action.description}`);
            console.log(`📋 [VersionRestorer] Action details:`, {
              id: action.id,
              type: action.operation?.op,
              affectedRanges: action.affectedRanges,
              beforeState: action.beforeState ? 'Present' : 'Missing'
            });
            
            await this.undoHandlers.undoAction(context, action);
            console.log(`✅ [VersionRestorer] Successfully undid action ${action.id}`);
            restoredCount++;
          } catch (error) {
            const errorMessage = `Error undoing action ${action.description}: ${error.message || error}`;
            console.error(`❌ [VersionRestorer] ${errorMessage}`);
            console.error(`❌ [VersionRestorer] Error stack:`, error.stack);
            errors.push(errorMessage);
          }
        }
        
        console.log(`🔄 [VersionRestorer] Syncing changes to Excel...`);
        try {
          await context.sync();
          console.log(`✅ [VersionRestorer] Context sync successful`);
        } catch (syncError) {
          console.error(`❌ [VersionRestorer] Error during context.sync():`, syncError);
          console.error(`❌ [VersionRestorer] Sync error stack:`, syncError.stack);
          errors.push(`Error syncing changes: ${syncError.message || syncError}`);
        }
        
        console.log(`🔄 [VersionRestorer] Version restore completed: ${restoredCount} actions restored`);
      });
      
      console.log(`✅ [VersionRestorer] Excel.run completed successfully`);
      const result = {
        success: errors.length === 0,
        restoredVersion: version,
        restoredActions: restoredCount,
        errors,
        restorePointId
      };
      
      console.log(`💾 [VersionRestorer] Returning result:`, result);
      return result;
    } catch (error) {
      const errorMessage = `Error restoring version: ${error.message || error}`;
      console.error(`❌ [VersionRestorer] ${errorMessage}`);
      console.error(`❌ [VersionRestorer] Error stack:`, error.stack);
      
      const result = {
        success: false,
        restoredVersion: version,
        restoredActions: restoredCount,
        errors: [errorMessage, ...errors],
        restorePointId
      };
      
      console.log(`💾 [VersionRestorer] Returning error result:`, result);
      return result;
    }
  }
}
