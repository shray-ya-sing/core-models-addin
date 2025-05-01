import * as React from 'react';
import { useState, useEffect } from 'react';

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

  // Define styles based on status
  const indicatorStyle: React.CSSProperties = {
    display: 'inline-block',
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    marginRight: '8px',
    verticalAlign: 'middle',
    backgroundColor: 
      status === StatusType.Pending ? '#f39c12' : // Yellow
      status === StatusType.Success ? '#2ecc71' : // Green
      status === StatusType.Error ? '#e74c3c' :   // Red
      '#95a5a6',                                  // Gray (idle)
    animation: status === StatusType.Pending ? 'pulse 1.5s infinite' : 'none',
    transition: 'background-color 0.3s ease-in-out',
  };

  const containerStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    padding: '8px 12px',
    marginBottom: '8px',
    fontSize: '12px',
    color: '#ffffff',
    backgroundColor: 'rgba(0, 0, 0, 0.5)',
    borderRadius: '4px',
    opacity: fadeOut ? 0 : 1,
    transition: 'opacity 0.5s ease-in-out',
  };

  return (
    <div style={containerStyle}>
      <div style={indicatorStyle} />
      <span>{message}</span>
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

export default StatusIndicator;
