/**
 * Models for pending changes approval system
 */

import { WorkbookAction, BeforeState, AffectedRange, VersionEventType } from './VersionModels';
import { ExcelOperation } from './ExcelOperationModels';

/**
 * Status of a pending change
 */
export enum PendingChangeStatus {
  PENDING = 'pending',
  ACCEPTED = 'accepted',
  REJECTED = 'rejected'
}

/**
 * A pending change that requires user approval
 */
export interface PendingChange {
  id: string;
  workbookId: string;
  operation: ExcelOperation;
  beforeState: BeforeState;
  status: PendingChangeStatus;
  timestamp: number;
  affectedRanges: string[];
  commandId?: string;
  description: string;
}

/**
 * Options for creating a pending change
 */
export interface CreatePendingChangeOptions {
  workbookId: string;
  operation: ExcelOperation;
  beforeState: BeforeState;
  commandId?: string;
  description?: string;
}

/**
 * Result of accepting or rejecting a pending change
 */
export interface PendingChangeResult {
  success: boolean;
  message: string;
  changeId: string;
}
