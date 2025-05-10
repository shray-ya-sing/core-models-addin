// src/client/components/chat/ApprovalModeToggle.tsx
import React from 'react';

interface ApprovalModeToggleProps {
  approvalEnabled: boolean;
  toggleApproval: () => void;
}

const ApprovalModeToggle: React.FC<ApprovalModeToggleProps> = ({ 
  approvalEnabled, 
  toggleApproval 
}) => {
  return (
    <div className="flex items-center justify-end px-2 py-0.5 border-b border-gray-800">
      <div 
        className="flex items-center cursor-pointer"
        onClick={toggleApproval}
      >
        <div 
          className="h-2 w-2 rounded-full mr-1"
          style={{ backgroundColor: approvalEnabled ? '#3b82f6' : '#6b7280' }}
        />
        <span 
          style={{ 
            fontSize: '8px', 
            color: approvalEnabled ? '#d1d5db' : '#9ca3af',
            fontFamily: 'Arial, "Helvetica Neue", Helvetica, sans-serif',
            letterSpacing: '-0.5px'
          }}
        >
          {approvalEnabled ? 'Approval Mode' : 'Auto Mode'}
        </span>
      </div>
    </div>
  );
};

export default ApprovalModeToggle;