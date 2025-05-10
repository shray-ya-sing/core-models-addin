import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import './VersionHistoryView.css';
import { VersionHistoryProvider } from '../services/versioning/VersionHistoryProvider';
import { WorkbookVersion, WorkbookAction, VersionEventType, RestoreVersionOptions, RestoreResult } from '../models/VersionModels';
import { ExcelOperationType } from '../models/ExcelOperationModels';
// No need to import VersionRestorer since we'll use the provider's restoreVersion method

// Define types for UI representation of version history
interface VersionChange {
  type: "added" | "removed" | "modified";
  path: string;
  changes: {
    before?: string;
    after?: string;
  };
}

interface VersionHistoryItem {
  id: string;
  type: "change" | "branch" | "merge" | "manual" | "auto";
  message: string;
  description: string; // More descriptive title for the version
  author: string;
  timestamp: number;
  branch?: string; // For backward compatibility with mock data
  actionIds: string[];
  changes: VersionChange[];
  // Icon and color information
  icon?: string;
  color?: string;
}

interface VersionHistoryViewProps {
  onClose: () => void;
  workbookId?: string;
  versionHistoryProvider?: VersionHistoryProvider;
}

const VersionHistoryView: React.FC<VersionHistoryViewProps> = ({ onClose, workbookId, versionHistoryProvider }) => {
  // State for version history data
  const [versionHistory, setVersionHistory] = useState<VersionHistoryItem[]>([]);
  
  // Selected version for detailed view
  const [selectedVersion, setSelectedVersion] = useState<string | null>(null);
  
  // Notification state
  const [notification, setNotification] = useState<{message: string, type: 'success' | 'error'} | null>(null);
  
  // No need for a separate version restorer state as we'll use the provider directly
  
  // Helper functions for icons and colors - memoized to prevent unnecessary re-renders
  const getIconForVersionType = useCallback((type: string): string => {
    switch (type) {
      case 'manual': return '💾';
      case 'auto': return '🔄';
      case 'branch': return '🔀';
      case 'merge': return '🔀';
      default: return '📝';
    }
  }, []);

  const getColorForVersionType = useCallback((type: string): string => {
    switch (type) {
      case 'manual': return '#4299e1'; // Blue
      case 'auto': return '#48bb78'; // Green
      case 'branch': return '#805ad5'; // Purple
      case 'merge': return '#d69e2e'; // Yellow
      default: return '#718096'; // Gray
    }
  }, []);

  const getIconForActionType = useCallback((type: string): string => {
    switch (type) {
      case 'set_value': return '✏️';
      case 'set_formula': return '🧮';
      case 'set_format': return '🎨';
      case 'add_worksheet': return '➕';
      case 'delete_worksheet': return '❌';
      case 'rename_worksheet': return '🏷️';
      default: return '📝';
    }
  }, []);

  const getColorForActionType = useCallback((type: string): string => {
    switch (type) {
      case 'set_value': return '#4299e1'; // Blue
      case 'set_formula': return '#48bb78'; // Green
      case 'set_format': return '#805ad5'; // Purple
      case 'add_worksheet': return '#d69e2e'; // Yellow
      case 'delete_worksheet': return '#f56565'; // Red
      case 'rename_worksheet': return '#ed8936'; // Orange
      default: return '#718096'; // Gray
    }
  }, []);

  // Helper function to format cell values for display - memoized
  const formatCellValue = useCallback((value: any): string => {
    if (value === null || value === undefined) return 'empty';
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  }, []);

  // Helper function to extract cell value from beforeState - memoized
  const extractCellValue = useCallback((action: WorkbookAction): string | null => {
    if (!action.beforeState || !action.beforeState.values) return null;
    
    // For set_value operations, extract the value from the beforeState
    if (action.beforeState.values.length > 0 && action.beforeState.values[0].length > 0) {
      return formatCellValue(action.beforeState.values[0][0]);
    }
    
    return null;
  }, [formatCellValue]);

  // Helper function to extract the new value from the operation - memoized
  const extractNewValue = useCallback((action: WorkbookAction): string | null => {
    if (!action.operation) return null;
    
    // For set_value operations, extract the new value
    if (action.operation.op === ExcelOperationType.SET_VALUE && 'value' in action.operation) {
      return formatCellValue(action.operation.value);
    }
    
    return null;
  }, [formatCellValue]);

  // Helper function to get cell reference from action - memoized
  const getCellReference = useCallback((action: WorkbookAction): string => {
    if (action.affectedRanges && action.affectedRanges.length > 0) {
      const range = action.affectedRanges[0];
      return `${range.sheetName}!${range.range || 'Unknown'}`;
    }
    return 'Unknown cell';
  }, []);
  
  // Generate a descriptive title based on version and actions - memoized
  const generateDescriptiveTitle = useCallback((version: WorkbookVersion, actions: WorkbookAction[]): string => {
    if (version.description) {
      return version.description;
    }
    
    if (actions.length === 0) {
      return 'Version created';
    }
    
    if (actions.length === 1) {
      return actions[0].description || `Operation: ${actions[0].type}`;
    }
    
    // Group actions by type
    const actionTypes = new Set(actions.map(a => a.type));
    if (actionTypes.size === 1) {
      return `${actions.length} ${Array.from(actionTypes)[0]} operations`;
    }
    
    return `${actions.length} operations (${Array.from(actionTypes).join(', ')})`;
  }, []);

  // Function to load version history - wrapped in useCallback to avoid dependency issues
  const loadVersionHistory = useCallback(() => {
    if (!versionHistoryProvider || !workbookId) {
      return;
    }
    
    try {
      console.log(`🔍 [VersionHistoryView] Loading version history for workbook: ${workbookId}`);
      
      // Get all versions
      const versions = versionHistoryProvider.getVersions();
      console.log(`Loaded ${versions.length} versions from provider`);
      
      // Create a combined history of versions and actions
      const historyItems: VersionHistoryItem[] = [];
      
      // First, add all versions to the history
      versions.forEach(v => {
        // Determine the type based on the version's event type
        let type: "change" | "branch" | "merge" | "manual" | "auto" = "change";
        if (v.type === VersionEventType.ManualSave) {
          type = "manual";
        } else if (v.type === VersionEventType.AutoSave) {
          type = "auto";
        }
        
        // Get the actions for this version
        const actions = versionHistoryProvider.getActionsForVersion(v.id) || [];
        
        // Convert actions to changes for UI display
        const changes: VersionChange[] = actions.flatMap(action => {
          // Create a change entry for each affected range
          return action.affectedRanges.map(range => ({
            type: "modified",
            path: `${range.sheetName}!${range.range || '[sheet-level]'}`,
            changes: {
              before: JSON.stringify(action.beforeState),
              after: "After state not captured"
            }
          }));
        });
        
        // Generate a descriptive title based on actions or default to version description
        const description = generateDescriptiveTitle(v, actions);
        
        // Add the version to the history items
        historyItems.push({
          id: v.id,
          type,
          message: v.description || 'Version created',
          description,
          author: v.author || 'System',
          timestamp: v.timestamp,
          actionIds: v.actionIds,
          changes,
          icon: getIconForVersionType(type),
          color: getColorForVersionType(type)
        });
      });
      
      // Now add any actions that aren't part of a version
      const versionActionIds = new Set<string>();
      versions.forEach(version => {
        version.actionIds.forEach(id => versionActionIds.add(id));
      });
      
      // Filter actions that aren't part of any version
      const allActions = versionHistoryProvider.getAllActions();
      const unversionedActions = allActions.filter(action => !versionActionIds.has(action.id));
      console.log(`Found ${unversionedActions.length} unversioned actions`);
      
      // Group unversioned actions by query processing boundaries
      const actionGroups: WorkbookAction[][] = [];
      const GROUPING_THRESHOLD_MS = 1000; // 1 second for tighter grouping
      
      // Sort actions by timestamp (newest first)
      unversionedActions.sort((a, b) => b.timestamp - a.timestamp);
      
      // Group actions by query processing boundaries
      unversionedActions.forEach(action => {
        const lastGroup = actionGroups[actionGroups.length - 1];
        
        // Check if this action belongs to the same query processing as the last group
        const isSameQuery = lastGroup && 
          // Check time proximity
          Math.abs(lastGroup[0].timestamp - action.timestamp) < GROUPING_THRESHOLD_MS &&
          // Check if they have the same query ID in metadata (if available)
          (action.metadata?.queryId && 
           lastGroup[0].metadata?.queryId && 
           action.metadata?.queryId === lastGroup[0].metadata?.queryId);
        
        if (isSameQuery) {
          // Add to existing group if from same query processing
          lastGroup.push(action);
        } else {
          // Create a new group
          actionGroups.push([action]);
        }
      });
      
      // Create history items for each action group
      actionGroups.forEach(group => {
        if (group.length === 0) return;
        
        // Use the timestamp of the first action in the group
        const timestamp = group[0].timestamp;
        
        // Create a unique ID for this group
        const groupId = `action-group-${timestamp}`;
        
        // Collect all action IDs in this group
        const actionIds = group.map(action => action.id);
        
        // Convert actions to changes for UI display
        const changes: VersionChange[] = group.flatMap(action => {
          return action.affectedRanges.map(range => ({
            type: "modified",
            path: `${range.sheetName}!${range.range || '[sheet-level]'}`,
            changes: {
              before: JSON.stringify(action.beforeState),
              after: "After state not captured"
            }
          }));
        });
        
        // Generate a descriptive message based on the actions
        const message = group.length === 1 
          ? group[0].description 
          : `${group.length} operations performed`;
        
        // Create a more detailed description
        const description = group.map(action => action.description).join(', ');
        
        // Add the action group to the history items
        historyItems.push({
          id: groupId,
          type: "change",
          message,
          description,
          author: 'System',
          timestamp,
          actionIds,
          changes,
          icon: getIconForActionType(group[0].type),
          color: getColorForActionType(group[0].type)
        });
      });
      
      // Sort all history items by timestamp (newest first)
      historyItems.sort((a, b) => b.timestamp - a.timestamp);
      
      // Set the version history
      setVersionHistory(historyItems);
    } catch (error) {
      console.error('Error loading version history:', error);
    }
  }, [versionHistoryProvider, workbookId, getIconForVersionType, getIconForActionType, getColorForVersionType, getColorForActionType, generateDescriptiveTitle]);
  
  // Load version history on component mount
  useEffect(() => {
    // Load history if we have both provider and workbook ID
    if (versionHistoryProvider && workbookId) {
      loadVersionHistory();
    }
  }, [versionHistoryProvider, workbookId, loadVersionHistory]);
  
  // Function to manually save a version
  const saveVersion = () => {
    if (!versionHistoryProvider) {
      console.error('Cannot save version: No version history provider available');
      setNotification({
        message: 'Cannot save version: Version history provider not available',
        type: 'error'
      });
      setTimeout(() => setNotification(null), 3000);
      return;
    }
    
    try {
      // Create a manual version point
      const versionId = versionHistoryProvider.createVersion({
        description: `Manual save at ${new Date().toLocaleString()}`,
        author: 'User',
        type: VersionEventType.ManualSave
      });
      
      // Show success message
      setNotification({
        message: 'Version saved successfully',
        type: 'success'
      });
      
      // Refresh the version history
      loadVersionHistory();
      
      // Clear notification after 3 seconds
      setTimeout(() => setNotification(null), 3000);
    } catch (error) {
      console.error('Error saving version:', error);
      setNotification({
        message: `Failed to save version: ${error instanceof Error ? error.message : String(error)}`,
        type: 'error'
      });
      
      // Clear notification after 5 seconds
      setTimeout(() => setNotification(null), 5000);
    }
  };

  // Function to clear all version history
  const clearVersionHistory = () => {
    if (!versionHistoryProvider) {
      console.error('Cannot clear version history: No version history provider available');
      setNotification({
        message: 'Cannot clear version history: Version history provider not available',
        type: 'error'
      });
      setTimeout(() => setNotification(null), 3000);
      return;
    }
    
    try {
      // Clear the version history
      versionHistoryProvider.clearVersionHistory();
      
      // Show success message
      setNotification({
        message: 'Version history cleared successfully',
        type: 'success'
      });
      
      // Refresh the version history (should be empty now)
      loadVersionHistory();
      
      // Clear notification after 3 seconds
      setTimeout(() => setNotification(null), 3000);
    } catch (error) {
      console.error('Error clearing version history:', error);
      setNotification({
        message: `Failed to clear version history: ${error instanceof Error ? error.message : String(error)}`,
        type: 'error'
      });
      
      // Clear notification after 5 seconds
      setTimeout(() => setNotification(null), 5000);
    }
  };
  
  // Function to restore a version
  const restoreVersion = async (versionId: string) => {
    console.log(`🔄 [VersionHistoryView] Restore version triggered for version ID: ${versionId}`);
    
    if (!versionHistoryProvider) {
      console.error('❌ [VersionHistoryView] Cannot restore version: No version history provider available');
      setNotification({
        message: 'Cannot restore version: Version history provider not available',
        type: 'error'
      });
      setTimeout(() => setNotification(null), 3000);
      return;
    }
    
    try {
      // Check if this is an action group (ID starts with 'action-group-')
      const isActionGroup = versionId.startsWith('action-group-');
      console.log(`📋 [VersionHistoryView] ID type: ${isActionGroup ? 'Action Group' : 'Formal Version'}`);
      
      // For action groups, we need to handle them differently
      if (isActionGroup) {
        // Find the version history item to get the action IDs
        const historyItem = versionHistory.find(item => item.id === versionId);
        
        if (!historyItem || !historyItem.actionIds || historyItem.actionIds.length === 0) {
          console.error(`❌ [VersionHistoryView] Cannot restore action group: No actions found for ID ${versionId}`);
          setNotification({
            message: 'Cannot restore: No actions found for this item',
            type: 'error'
          });
          setTimeout(() => setNotification(null), 3000);
          return;
        }
        
        console.log(`📋 [VersionHistoryView] Found action group with ${historyItem.actionIds.length} actions:`, historyItem.actionIds);
        
        // Get the actual actions from the action IDs
        const allActions = versionHistoryProvider.getAllActions();
        const actionsToRestore = historyItem.actionIds
          .map(id => allActions.find(action => action.id === id))
          .filter(a => a !== undefined) as WorkbookAction[];
        
        console.log(`📋 [VersionHistoryView] Resolved ${actionsToRestore.length} actions to restore:`, 
          actionsToRestore.map(a => ({ id: a.id, type: a.operation?.op, description: a.description })));
        
        if (actionsToRestore.length === 0) {
          console.error(`❌ [VersionHistoryView] Cannot restore action group: No valid actions found`);
          setNotification({
            message: 'Cannot restore: No valid actions found',
            type: 'error'
          });
          setTimeout(() => setNotification(null), 3000);
          return;
        }
        
        // Create a temporary version for the action group
        console.log(`🔄 [VersionHistoryView] Creating temporary version for action group`);
        
        // First, create a basic version
        const tempVersionId = versionHistoryProvider.createVersion({
          description: `Temporary version for restoring action group ${versionId}`,
          type: VersionEventType.Restore,
          author: 'System'
        });
        
        // Then manually associate the actions with this version
        // This is a workaround since CreateVersionOptions doesn't have actionIds
        console.log(`💾 [VersionHistoryView] Created temporary version with ID: ${tempVersionId}`);
        console.log(`💾 [VersionHistoryView] Manually associating ${actionsToRestore.length} actions with the temporary version`);
        
        // We need to get the version we just created and update it
        const tempVersion = versionHistoryProvider.getVersion(tempVersionId);
        if (tempVersion) {
          // Update the version's actionIds directly through the VersionHistoryService
          // This is implementation-specific and may need to be adjusted
          try {
            // We're assuming the VersionHistoryService has a way to update a version's actionIds
            // If not, we'll need to modify the VersionHistoryService to support this
            const allVersions = JSON.parse(localStorage.getItem(`version-history-versions-${tempVersion.workbookId}`) || '[]');
            const versionIndex = allVersions.findIndex((v: any) => v.id === tempVersionId);
            
            if (versionIndex >= 0) {
              allVersions[versionIndex].actionIds = actionsToRestore.map(a => a.id);
              localStorage.setItem(`version-history-versions-${tempVersion.workbookId}`, JSON.stringify(allVersions));
              console.log(`✅ [VersionHistoryView] Successfully updated temporary version with action IDs`);
            } else {
              console.error(`❌ [VersionHistoryView] Could not find temporary version in localStorage`);
            }
          } catch (err) {
            console.error(`❌ [VersionHistoryView] Error updating temporary version:`, err);
            console.error('❌ [VersionHistoryView] Error stack:', err instanceof Error ? err.stack : '');
          }
        } else {
          console.error(`❌ [VersionHistoryView] Could not retrieve the created temporary version`);
        }
        
        // Now restore using this temporary version
        const options: RestoreVersionOptions = {
          versionId: tempVersionId,
          createRestorePoint: true,
          restorePointDescription: `Restore point before reverting to ${tempVersion?.description}`
        };
        
        console.log(`🔄 [VersionHistoryView] Calling restoreVersion with options:`, options);
        
        const result = await versionHistoryProvider.restoreVersion(options);
        
        // Log the result
        console.log(`✅ [VersionHistoryView] Action group restore completed with result:`, result);
        
        // Show success message
        setNotification({
          message: `Action group ${versionId.substring(0, 15)}... restored successfully`,
          type: 'success'
        });
        setTimeout(() => setNotification(null), 3000);
        
        // Refresh the version history to show the new temporary version
        loadVersionHistory();
      } else {
        // This is a formal version, handle it normally
        // Log the version details
        const version = versionHistoryProvider.getVersion(versionId);
        console.log(`📋 [VersionHistoryView] Restoring formal version:`, version);
        
        // Log the actions associated with this version
        const actions = versionHistoryProvider.getActionsForVersion(versionId);
        console.log(`📋 [VersionHistoryView] Version has ${actions.length} actions to restore:`, 
          actions.map(a => ({ id: a.id, type: a.operation?.op, description: a.description })));
        
        // Create restore options
        const options: RestoreVersionOptions = {
          versionId: versionId,
          createRestorePoint: true,
          restorePointDescription: `Restore point before reverting to ${version?.description}`
        };
        
        console.log(`🔄 [VersionHistoryView] Calling restoreVersion with options:`, options);
        
        const result = await versionHistoryProvider.restoreVersion(options);
        
        // Log the result
        console.log(`✅ [VersionHistoryView] Version restore completed with result:`, result);
        
        // Show success message
        setNotification({
          message: `Version ${versionId.substring(0, 8)}... restored successfully`,
          type: 'success'
        });
        setTimeout(() => setNotification(null), 3000);
        
        // Refresh the version history
        loadVersionHistory();
      }
    } catch (err) {
      console.error('❌ [VersionHistoryView] Error restoring version:', err);
      console.error('❌ [VersionHistoryView] Error stack:', err instanceof Error ? err.stack : '');
      
      setNotification({
        message: `Failed to restore version: ${err instanceof Error ? err.message : String(err)}`,
        type: 'error'
      });
      setTimeout(() => setNotification(null), 5000);
    }
  };
  
  // Function to restore a single action
  const restoreSingleAction = async (actionId: string) => {
    console.log(`🔄 [VersionHistoryView] Restore single action triggered for action ID: ${actionId}`);
    
    if (!versionHistoryProvider) {
      console.error('❌ [VersionHistoryView] Cannot restore action: No version history provider available');
      setNotification({
        message: 'Cannot restore action: Version history provider not available',
        type: 'error'
      });
      setTimeout(() => setNotification(null), 3000);
      return;
    }
    
    try {
      // Get the action
      const action = versionHistoryProvider.getAllActions().find(action => action.id === actionId);
      
      if (!action) {
        console.error(`❌ [VersionHistoryView] Action not found: ${actionId}`);
        setNotification({
          message: 'Action not found',
          type: 'error'
        });
        setTimeout(() => setNotification(null), 3000);
        return;
      }
      
      console.log(`📋 [VersionHistoryView] Restoring single action:`, action);
      
      // Create a temporary version with just this action
      const tempVersionId = versionHistoryProvider.createVersion({
        description: `Restore single action: ${getCellReference(action)} - ${extractCellValue(action)}`,
        type: VersionEventType.Restore,
        author: 'System'
      });
      
      // Manually associate the action with this version
      console.log(`💾 [VersionHistoryView] Created temporary version with ID: ${tempVersionId}`);
      
      // Update the version's actionIds directly
      const tempVersion = versionHistoryProvider.getVersion(tempVersionId);
      if (tempVersion) {
        try {
          const allVersions = JSON.parse(localStorage.getItem(`version-history-versions-${tempVersion.workbookId}`) || '[]');
          const versionIndex = allVersions.findIndex((v: any) => v.id === tempVersionId);
          
          if (versionIndex >= 0) {
            allVersions[versionIndex].actionIds = [actionId];
            localStorage.setItem(`version-history-versions-${tempVersion.workbookId}`, JSON.stringify(allVersions));
            console.log(`✅ [VersionHistoryView] Successfully updated temporary version with action ID`);
          } else {
            console.error(`❌ [VersionHistoryView] Could not find temporary version in localStorage`);
          }
        } catch (err) {
          console.error(`❌ [VersionHistoryView] Error updating temporary version:`, err);
        }
      }
      
      // Restore the version
      const options: RestoreVersionOptions = {
        versionId: tempVersionId,
        createRestorePoint: true,
        restorePointDescription: `Restore point before reverting single action: ${getCellReference(action)}`
      };
      
      const result = await versionHistoryProvider.restoreVersion(options);
      
      if (result.success) {
        console.log(`✅ [VersionHistoryView] Successfully restored single action`);
        setNotification({
          message: `Successfully restored ${getCellReference(action)}`,
          type: 'success'
        });
        
        // Refresh the version history
        loadVersionHistory();
      } else {
        console.error(`❌ [VersionHistoryView] Failed to restore single action:`, result.errors);
        setNotification({
          message: `Failed to restore action: ${result.errors?.join(', ')}`,
          type: 'error'
        });
      }
      
      // Clear notification after 3 seconds
      setTimeout(() => setNotification(null), 3000);
    } catch (err) {
      console.error(`❌ [VersionHistoryView] Error during single action restoration:`, err);
      setNotification({
        message: `Error during restoration: ${err instanceof Error ? err.message : String(err)}`,
        type: 'error'
      });
      setTimeout(() => setNotification(null), 5000);
    }
  };

  // Render the component
  return (
    <div className="version-history-view">
      <div className="version-history-header">
        <h2>Version History</h2>
        <button className="close-button" onClick={onClose}>×</button>
      </div>
      
      <div className="version-history-actions">
        <button className="save-button" onClick={saveVersion}>Save Version</button>
        <button className="clear-button" onClick={clearVersionHistory}>Clear History</button>
      </div>
      
      {notification && (
        <div className={`notification ${notification.type}`}>
          {notification.message}
        </div>
      )}
      
      <div className="version-history-list">
        {versionHistory.length === 0 ? (
          <div className="empty-state">No version history available</div>
        ) : (
          versionHistory.map(item => {
            // Find the actions associated with this version item
            const actions = item.actionIds.map(id => 
              versionHistoryProvider?.getAllActions().find(action => action.id === id)
            ).filter(a => a !== undefined) as WorkbookAction[];
            
            return (
              <div key={item.id} className="version-group">
                <div 
                  className={`version-header ${selectedVersion === item.id ? 'expanded' : ''}`}
                  onClick={() => setSelectedVersion(item.id === selectedVersion ? null : item.id)}
                >
                  <div className="version-summary">
                    <div className="version-title">{item.message}</div>
                    <div className="version-meta">
                      {item.author} • {new Date(item.timestamp).toLocaleString()}
                    </div>
                  </div>
                  <div className="version-indicator">
                    {selectedVersion === item.id ? '▼' : '▶'}
                  </div>
                </div>
                
                {selectedVersion === item.id && (
                  <div className="version-details">
                    <div className="version-changes">
                      {actions.map(action => {
                        const oldValue = extractCellValue(action);
                        const newValue = extractNewValue(action);
                        const cellRef = getCellReference(action);
                        
                        return (
                          <div key={action.id} className="change-item">
                            <div className="change-header">
                              <div className="change-description">
                                {action.description}
                              </div>
                              <div className="change-actions">
                                <button 
                                  className="action-button"
                                  onClick={(e) => {
                                    e.stopPropagation(); // Prevent expanding/collapsing the parent
                                    restoreSingleAction(action.id);
                                  }}
                                  title="Restore just this action"
                                >
                                  Restore
                                </button>
                              </div>
                            </div>
                            
                            {oldValue && oldValue !== 'empty' && (
                              <div className="change-value removed">
                                - {cellRef}: {oldValue}
                              </div>
                            )}
                            
                            {newValue && (
                              <div className="change-value added">
                                + {cellRef}: {newValue}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    
                    <div className="version-actions">
                      <button 
                        className="view-button"
                        onClick={() => console.log(`View at point: ${item.id}`)}
                      >
                        View Model at This Point
                      </button>
                      <button 
                        className="restore-button"
                        onClick={() => restoreVersion(item.id)}
                      >
                        Restore This Version
                      </button>
                    </div>
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}
    
export { VersionHistoryView };