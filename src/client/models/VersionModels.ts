/**
 * Models for the version history tracking system
 */

import { ExcelOperation } from './ExcelOperationModels';

/**
 * Types of version events that can be recorded
 */
export enum VersionEventType {
  // User-initiated events
  ManualSave = 'manualSave',
  Restore = 'restore',
  
  // System events
  AutoSave = 'autoSave',
  InitialState = 'initialState',
  
  // Operation types
  CellOperation = 'cellOperation',
  RangeOperation = 'rangeOperation',
  SheetOperation = 'sheetOperation',
  WorkbookOperation = 'workbookOperation',
  FormatOperation = 'formatOperation',
  ChartOperation = 'chartOperation',
  TableOperation = 'tableOperation',
  CompositeOperation = 'compositeOperation'
}

/**
 * Represents a single recorded action on the workbook
 */
export interface WorkbookAction {
  id: string;                     // Unique identifier for this action
  workbookId: string;             // ID of the workbook this action belongs to
  timestamp: number;              // When the action occurred
  type: VersionEventType;         // Type of action
  operation: ExcelOperation;      // The Excel operation that was performed
  description: string;            // Human-readable description of the action
  affectedRanges: AffectedRange[]; // Ranges affected by this action
  beforeState: BeforeState;       // State before the action was performed
  metadata?: Record<string, any>; // Additional metadata about the action
}

/**
 * Represents a range affected by an action
 */
export interface AffectedRange {
  sheetName: string;              // Name of the sheet
  range: string;                  // Range in A1 notation (e.g., "A1:B10")
  type: 'cell' | 'range' | 'sheet' | 'table' | 'chart'; // Type of affected entity
}

/**
 * Represents the state before an action was performed
 */
export interface BeforeState {
  values?: any[][];               // Cell values before the change
  formulas?: string[][];          // Formulas before the change
  formats?: Record<string, any>[]; // Formatting before the change
  sheetProperties?: Record<string, any>; // Sheet properties before the change
  otherState?: Record<string, any>; // Any other state that needs to be preserved
}

/**
 * Represents a version point in the workbook history
 */
export interface WorkbookVersion {
  id: string;                     // Unique identifier for this version
  workbookId: string;             // ID of the workbook this version belongs to
  timestamp: number;              // When this version was created
  type: VersionEventType;         // Type of version (manual, auto, etc.)
  description: string;            // User-provided or auto-generated description
  author: string;                 // Who created this version
  actionIds: string[];            // IDs of actions included in this version
  tags?: string[];                // Optional tags for categorization
}

/**
 * Represents a group of related actions that form a logical version
 */
export interface VersionGroup {
  version: WorkbookVersion;       // Version metadata
  actions: WorkbookAction[];      // The actions in this version
}

/**
 * Options for creating a new version
 */
export interface CreateVersionOptions {
  description?: string;           // Optional description
  type?: VersionEventType;        // Type of version (defaults to ManualSave)
  author?: string;                // Who is creating the version
  tags?: string[];                // Optional tags
  includeActionsSince?: number;   // Timestamp to include actions from
}

/**
 * Options for restoring a version
 */
export interface RestoreVersionOptions {
  versionId: string;              // ID of the version to restore
  createRestorePoint?: boolean;   // Whether to create a restore point before restoring
  restorePointDescription?: string; // Description for the restore point
  selective?: boolean;            // Whether to selectively restore parts of the version
  selectiveRanges?: AffectedRange[]; // Ranges to selectively restore (if selective is true)
}

/**
 * Result of a version restoration operation
 */
export interface RestoreResult {
  success: boolean;               // Whether the restore was successful
  restoredVersion: WorkbookVersion; // The version that was restored
  restoredActions: number;        // Number of actions that were restored
  errors?: string[];              // Any errors that occurred during restoration
  restorePointId?: string;        // ID of the restore point created before restoration
}
