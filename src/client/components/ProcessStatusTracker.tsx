import * as React from 'react';
import { useState, useEffect } from 'react';
import { 
  ProcessStage, 
  ProcessStatus, 
  ProcessStatusEvent, 
  ProcessStatusManager 
} from '../models/ProcessStatusModels';
import { StatusType } from './StatusIndicator';

/**
 * Maps ProcessStatus to StatusType
 */
function mapProcessStatusToStatusType(status: ProcessStatus): StatusType {
  switch (status) {
    case ProcessStatus.Pending:
      return StatusType.Pending;
    case ProcessStatus.Success:
      return StatusType.Success;
    case ProcessStatus.Error:
      return StatusType.Error;
    default:
      return StatusType.Idle;
  }
}

/**
 * Gets a human-readable name for a process stage
 */
export function getStageName(stage: ProcessStage): string {
  switch (stage) {
    case ProcessStage.KnowledgeBaseQuery:
      return 'Knowledge Base Search';
    case ProcessStage.WorkbookCapture:
      return 'Excel Workbook Capture';
    case ProcessStage.QueryProcessing:
      return 'Query Processing';
    case ProcessStage.CommandPlanning:
      return 'Command Planning';
    case ProcessStage.CommandExecution:
      return 'Command Execution';
    case ProcessStage.OperationExecution:
      return 'Operation Execution';
    default:
      return 'Unknown Stage';
  }
}

interface ProcessStatusItemProps {
  event: ProcessStatusEvent;
}

/**
 * Component to display a single process status item
 */
const ProcessStatusItem: React.FC<ProcessStatusItemProps> = ({ event }) => {
  const indicatorStyle: React.CSSProperties = {
    display: 'inline-block',
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    marginRight: '8px',
    verticalAlign: 'middle',
    backgroundColor: 
      event.status === ProcessStatus.Pending ? '#f39c12' : // Yellow
      event.status === ProcessStatus.Success ? '#2ecc71' : // Green
      event.status === ProcessStatus.Error ? '#e74c3c' :   // Red
      '#95a5a6',                                  // Gray (idle)
    animation: event.status === ProcessStatus.Pending ? 'pulse 1.5s infinite' : 'none',
    transition: 'background-color 0.3s ease-in-out',
  };

  const containerStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    padding: '6px 12px',
    marginBottom: '4px',
    fontSize: '12px',
    color: '#ffffff',
    backgroundColor: 'rgba(0, 0, 0, 0.5)',
    borderRadius: '4px',
  };

  const stageLabelStyle: React.CSSProperties = {
    fontWeight: 'bold',
    marginRight: '8px',
  };
  
  return (
    <div style={containerStyle}>
      <div style={indicatorStyle} />
      <span style={stageLabelStyle}>{getStageName(event.stage)}:</span>
      <span>{event.message}</span>
    </div>
  );
};

interface ProcessStatusTrackerProps {
  processId?: string; // If provided, shows only statuses for this process ID
  maxItems?: number;  // Maximum number of items to show (default: all)
}

/**
 * Component that tracks and displays all process statuses in the application
 */
const ProcessStatusTracker: React.FC<ProcessStatusTrackerProps> = ({ 
  processId,
  maxItems = 10
}) => {
  const [statusEvents, setStatusEvents] = useState<ProcessStatusEvent[]>([]);
  
  useEffect(() => {
    const processManager = ProcessStatusManager.getInstance();
    
    // Initial load of existing statuses
    if (processId) {
      setStatusEvents(processManager.getStatusesForProcess(processId));
    }
    
    // Listen for new status events
    const removeListener = processManager.addListener((event) => {
      if (!processId || event.id === processId) {
        setStatusEvents(prev => {
          // Add the new event at the end
          const newEvents = [...prev, event];
          
          // Filter out older events of the same stage/process to avoid duplicates
          const filtered = newEvents.filter((e, index, self) => {
            // Keep this event if it's the latest for its process+stage combination
            return index === self.findIndex(other => 
              other.id === e.id && other.stage === e.stage && other.timestamp >= e.timestamp
            );
          });
          
          // Sort by timestamp (newest first)
          const sorted = filtered.sort((a, b) => b.timestamp - a.timestamp);
          
          // Limit the number of items if needed
          return sorted.slice(0, maxItems);
        });
      }
    });
    
    return removeListener;
  }, [processId, maxItems]);
  
  // Sort by timestamp (newest first)
  const sortedEvents = [...statusEvents].sort((a, b) => b.timestamp - a.timestamp);
  
  return (
    <div style={{ marginBottom: '10px' }}>
      {sortedEvents.map((event, index) => (
        <ProcessStatusItem key={`${event.id}_${event.stage}_${index}`} event={event} />
      ))}
      <style>
        {`
        @keyframes pulse {
          0% { opacity: 1; }
          50% { opacity: 0.4; }
          100% { opacity: 1; }
        }
        `}
      </style>
    </div>
  );
};

export default ProcessStatusTracker;
