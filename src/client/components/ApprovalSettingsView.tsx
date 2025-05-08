/**
 * ApprovalSettingsView Component
 * 
 * Provides UI controls to enable/disable the approval workflow for AI-generated changes.
 */

import * as React from 'react';
import { ClientExcelCommandInterpreter } from '../services/ClientExcelCommandInterpreter';
import { PendingChangesTracker } from '../services/PendingChangesTracker';

interface ApprovalSettingsViewProps {
  commandInterpreter: ClientExcelCommandInterpreter;
  pendingChangesTracker: PendingChangesTracker;
}

interface ApprovalSettingsViewState {
  approvalEnabled: boolean;
  pendingChangesCount: number;
}

/**
 * Component for controlling the AI changes approval workflow
 */
export class ApprovalSettingsView extends React.Component<ApprovalSettingsViewProps, ApprovalSettingsViewState> {
  private refreshInterval: number | null = null;
  
  constructor(props: ApprovalSettingsViewProps) {
    super(props);
    
    this.state = {
      approvalEnabled: false,
      pendingChangesCount: 0
    };
  }
  
  /**
   * Component mounted
   */
  public componentDidMount(): void {
    // Start polling for pending changes count
    this.refreshInterval = window.setInterval(() => this.refreshPendingChangesCount(), 1000);
  }
  
  /**
   * Component will unmount
   */
  public componentWillUnmount(): void {
    // Stop polling
    if (this.refreshInterval !== null) {
      window.clearInterval(this.refreshInterval);
      this.refreshInterval = null;
    }
  }
  
  /**
   * Refresh the pending changes count
   */
  private refreshPendingChangesCount(): void {
    // Get the current workbook ID from the command interpreter
    const workbookId = this.props.commandInterpreter.getCurrentWorkbookId();
    
    if (!workbookId) {
      return;
    }
    
    // Get pending changes for the current workbook
    const pendingChanges = this.props.pendingChangesTracker.getPendingChanges(workbookId);
    
    // Update state if count has changed
    if (pendingChanges.length !== this.state.pendingChangesCount) {
      this.setState({ pendingChangesCount: pendingChanges.length });
    }
  }
  
  /**
   * Toggle the approval workflow
   */
  private toggleApprovalWorkflow = (): void => {
    const newState = !this.state.approvalEnabled;
    
    // Update the command interpreter
    this.props.commandInterpreter.setRequireApproval(newState);
    
    // Update state
    this.setState({ approvalEnabled: newState });
  };
  
  /**
   * Render the component
   */
  public render(): JSX.Element {
    const { approvalEnabled, pendingChangesCount } = this.state;
    
    return (
      <div className="approval-settings p-4 border-t border-gray-200 mt-4">
        <h3 className="text-sm font-mono mb-2 text-gray-600">AI Changes Approval</h3>
        
        <div className="flex items-center justify-between">
          <div className="flex items-center">
            <div 
              className={`w-10 h-5 flex items-center rounded-full p-1 cursor-pointer ${
                approvalEnabled ? 'bg-green-500' : 'bg-gray-300'
              }`}
              onClick={this.toggleApprovalWorkflow}
            >
              <div
                className={`bg-white w-4 h-4 rounded-full shadow-md transform transition-transform duration-300 ${
                  approvalEnabled ? 'translate-x-5' : 'translate-x-0'
                }`}
              />
            </div>
            <span className="ml-2 text-xs font-mono text-gray-600">
              {approvalEnabled ? 'Enabled' : 'Disabled'}
            </span>
          </div>
          
          {pendingChangesCount > 0 && (
            <div className="bg-green-100 text-green-800 text-xs font-mono px-2 py-1 rounded">
              {pendingChangesCount} pending change{pendingChangesCount !== 1 ? 's' : ''}
            </div>
          )}
        </div>
        
        <p className="text-xs font-mono text-gray-500 mt-2">
          {approvalEnabled 
            ? 'AI-generated changes will be highlighted in green with accept/reject buttons'
            : 'AI-generated changes will be applied immediately'}
        </p>
      </div>
    );
  }
}
