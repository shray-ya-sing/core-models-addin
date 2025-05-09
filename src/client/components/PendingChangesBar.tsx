/**
 * PendingChangesBar Component
 * 
 * Displays a bar with pending changes and accept/reject buttons
 */

import * as React from 'react';
import { PendingChange, PendingChangeStatus } from '../services/PendingChangesTracker';

interface PendingChangesBarProps {
  pendingChanges: PendingChange[];
  onAcceptAll: () => void;
  onRejectAll: () => void;
}

/**
 * Component for displaying pending changes with accept/reject buttons
 */
export const PendingChangesBar: React.FC<PendingChangesBarProps> = ({
  pendingChanges,
  onAcceptAll,
  onRejectAll
}) => {
  if (pendingChanges.length === 0) {
    return null;
  }

  return (
    <div className="bg-gray-800 border-t border-gray-700 px-4 py-2">
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center">
          <svg 
            className="w-4 h-4 text-green-500 mr-2" 
            fill="none" 
            stroke="currentColor" 
            viewBox="0 0 24 24" 
            xmlns="http://www.w3.org/2000/svg"
          >
            <path 
              strokeLinecap="round" 
              strokeLinejoin="round" 
              strokeWidth="2" 
              d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
            />
          </svg>
          <span style={{ 
            fontSize: 'clamp(10px, 1.25vw, 12px)', 
            fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
            color: '#d1d5db'
          }}>
            {pendingChanges.length} {pendingChanges.length === 1 ? 'change' : 'changes'} pending
          </span>
        </div>
        <div className="flex items-center space-x-2">
          <button
            onClick={onRejectAll}
            style={{ 
              fontSize: 'clamp(9px, 1vw, 11px)', 
              fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
              color: '#d1d5db',
              padding: '4px 8px',
              borderRadius: '4px',
              transition: 'all 0.2s'
            }}
            className="hover:text-white hover:bg-gray-700"
          >
            Reject all
          </button>
          <button
            onClick={onAcceptAll}
            style={{ 
              fontSize: 'clamp(9px, 1vw, 11px)', 
              fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
              backgroundColor: '#2563eb',
              color: 'white',
              padding: '4px 8px',
              borderRadius: '4px',
              transition: 'all 0.2s'
            }}
            className="hover:bg-blue-700"
          >
            Accept all
          </button>
        </div>
      </div>
      
      {/* No individual changes listed - minimalist UI */}
    </div>
  );
};

export default PendingChangesBar;
