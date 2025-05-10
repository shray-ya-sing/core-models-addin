/**
 * AI Approval System
 * 
 * Factory for creating and initializing the AI changes approval system.
 * This system allows users to review, accept, or reject changes made by AI.
 */

import { ClientExcelCommandInterpreter } from '../actions/ClientExcelCommandInterpreter';
import { VersionHistoryService } from '../versioning/VersionHistoryService';
import { ActionRecorder } from '../versioning/ActionRecorder';
import { PendingChangesTracker } from './PendingChangesTracker';
import { ShapeEventHandler } from '../ShapeEventHandler';

/**
 * Factory for creating and initializing the AI changes approval system
 */
export class AIApprovalSystem {
  /**
   * Initialize the AI changes approval system
   * @param commandInterpreter The Excel command interpreter
   * @param versionHistoryService The version history service
   * @returns An object containing the initialized components
   */
  public static initialize(
    commandInterpreter: ClientExcelCommandInterpreter,
    versionHistoryService: VersionHistoryService
  ): {
    pendingChangesTracker: PendingChangesTracker;
    shapeEventHandler: ShapeEventHandler;
  } {
    // Create the action recorder if it doesn't exist
    let actionRecorder = commandInterpreter.getActionRecorder();
    if (!actionRecorder) {
      actionRecorder = new ActionRecorder(versionHistoryService);
      commandInterpreter.setActionRecorder(actionRecorder);
    }
    
    // Create the pending changes tracker
    const pendingChangesTracker = new PendingChangesTracker(versionHistoryService, actionRecorder);
    
    // Create the shape event handler
    const shapeEventHandler = new ShapeEventHandler(pendingChangesTracker);
    
    // Set up the command interpreter with the approval system
    commandInterpreter.setPendingChangesTracker(pendingChangesTracker, shapeEventHandler);
    
    // By default, approval is disabled
    commandInterpreter.setRequireApproval(false);
    
    console.log('✅ AI changes approval system initialized');
    
    return {
      pendingChangesTracker,
      shapeEventHandler
    };
  }
}
