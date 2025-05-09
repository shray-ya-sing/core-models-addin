import * as React from 'react';
import { useState, useEffect } from 'react';
import { FileIcon, CheckIcon, AlertTriangleIcon } from '../components/icons';
import { cn } from '../utils/classUtils';

export enum StatusType {
  Pending = 'pending',
  Success = 'success',
  Error = 'error',
  Idle = 'idle'
}

interface StatusIndicatorProps {
  status: StatusType;
  message: string;
  autoHide?: boolean;
  hideAfterMs?: number;
}

/**
 * A status indicator component that shows a colored circle with a message
 * - Yellow pulsing circle for pending status
 * - Green circle for success status
 * - Red circle for error status
 */
export const StatusIndicator: React.FC<StatusIndicatorProps> = ({ 
  status, 
  message,
  autoHide = false,
  hideAfterMs = 5000
}) => {
  const [visible, setVisible] = useState(true);
  const [fadeOut, setFadeOut] = useState(false);

  useEffect(() => {
    // Reset visibility when status changes
    setVisible(true);
    setFadeOut(false);
    
    if (autoHide && (status === StatusType.Success || status === StatusType.Error)) {
      // Start fade out animation after a delay
      const fadeTimeout = setTimeout(() => {
        setFadeOut(true);
      }, hideAfterMs - 500); // Start fade 500ms before hiding
      
      // Hide component after specified delay
      const hideTimeout = setTimeout(() => {
        setVisible(false);
      }, hideAfterMs);
      
      // Clean up timeouts
      return () => {
        clearTimeout(fadeTimeout);
        clearTimeout(hideTimeout);
      };
    }
    
    // Return an empty cleanup function for code paths that don't set timeouts
    return () => {};
  }, [status, autoHide, hideAfterMs]);

  // Don't render if not visible
  if (!visible) return null;

  // Status-dependent classes
  const getStatusClasses = () => {
    switch(status) {
      case StatusType.Pending:
      case StatusType.Idle:
        return 'text-white';
      case StatusType.Success:
        return 'text-green-400';
      case StatusType.Error:
        return 'text-red-400';
      default:
        return 'text-gray-400';
    }
  };

  // Define a traffic light that will be visible for all statuses
  const getTrafficLight = () => {
    switch(status) {
      case StatusType.Pending:
      case StatusType.Idle:
        return <div className="w-4 h-4 rounded-full bg-yellow-500 border border-yellow-400 animate-pulse" />;
      case StatusType.Success:
        return <div className="w-4 h-4 rounded-full bg-green-500 border border-green-400" />;
      case StatusType.Error:
        return <div className="w-4 h-4 rounded-full bg-red-500 border border-red-400" />;
      default:
        return <div className="w-4 h-4 rounded-full bg-gray-500 border border-gray-400" />;
    }
  };

  // For debugging - log the current status
  console.log('StatusIndicator rendering with status:', status, 'and message:', message);

  return (
    <div className={cn(
      'flex items-center gap-3 rounded-md px-3 py-2 shadow-md bg-gray-800/70 border border-gray-700',
      fadeOut ? 'opacity-0' : 'opacity-100',
      'transition-opacity duration-500 ease-in-out'
    )}>
      <div className="flex items-center gap-3">
        {/* Debug traffic light - always grey regardless of status */}
        <div className="w-4 h-4 rounded-full bg-gray-400 border border-gray-300" />
        
        {/* Actual traffic light based on status */}
        {getTrafficLight()}
        
        {/* White text with smaller font size */}
        <span className="text-white" style={{ 
          fontSize: 'clamp(10px, 1.25vw, 12px)', 
          fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace' 
        }}>
          {message} [Status: {status}]
        </span>
      </div>
    </div>
  );
};

export default StatusIndicator;
