/**
 * Status of a command
 */
export enum CommandStatus {
  Pending = 'pending',
  Running = 'running',
  Completed = 'completed',
  Failed = 'failed'
}

/**
 * Types of queries that can be processed
 */
export enum QueryType {
  Greeting = 'greeting',
  WorkbookQuestion = 'workbookQuestion',
  WorkbookQuestionWithKB = 'workbookQuestionWithKB',
  WorkbookCommand = 'workbookCommand',
  WorkbookCommandWithKB = 'workbookCommandWithKB',
  Unknown = 'unknown'
}

/**
 * Operation type
 */
export enum OperationType {
  SetValue = 'setValue',
  SetFormula = 'setFormula',
  FormatCell = 'formatCell',
  CreateSheet = 'createSheet',
  DeleteSheet = 'deleteSheet',
  RenameSheet = 'renameSheet',
  CreateTable = 'createTable',
  CreateChart = 'createChart'
}

/**
 * Operation interface
 */
export interface Operation {
  type: OperationType;
  target: string;
  value?: any;
  options?: any;
  status?: 'pending' | 'running' | 'completed' | 'failed';
  error?: string;
}

/**
 * Command step interface
 */
export interface CommandStep {
  description: string;
  operations: Operation[];
  status?: 'pending' | 'running' | 'completed' | 'failed';
  error?: string;
}

/**
 * Command interface
 */
export interface Command {
  id: string;
  description: string;
  steps: CommandStep[];
  status: CommandStatus;
  progress?: number;
  error?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

/**
 * Workbook state interface
 */
export interface WorkbookState {
  sheets: SheetState[];
  activeSheet?: string;
}

/**
 * Sheet state interface
 */
export interface SheetState {
  name: string;
  values?: any[][];
  formulas?: any[][];
  formats?: any[][];
  tables?: any[];
  charts?: any[];
  namedRanges?: {name: string, value: string}[]; 
  usedRange?: {
    rowCount: number;
    columnCount: number;
  };
}

/**
 * Compressed workbook interface
 */
export interface CompressedWorkbook {
  sheets: CompressedSheet[];
  activeSheet?: string;
  metrics?: {
    totalSheets: number;
    totalCells: number;
    totalFormulas: number;
    totalTables: number;
    totalCharts: number;
  };
  dependencyGraph?: any;
  colorLegend?: any[];
  modelType?: string;
}

/**
 * Compressed sheet interface
 */
export interface CompressedSheet {
  name: string;
  summary?: string;  // Sheet summary description
  keyRegions?: {
    name: string;
    range: string;
    description?: string;
  }[];
  anchors?: {
    cell: string;
    value: any;
    type: string;
  }[];
  tables?: {
    name: string;
    range: string;
    headers: string[];
  }[];
  charts?: {
    name: string;
    type: string;
    range: string;
  }[];
  formulas?: any[][]; // 2D array of formulas (parallel to values) used for dependency analysis.
  metrics?: {
    rowCount: number;
    columnCount: number;
    formulaCount: number;
    valueCount: number;
    emptyCount: number;
  };
  cells?: {
    address: string;
    value: any;
    formula?: string;
    type: string;
    format?: any;
  }[];
}

/**
 * Workbook metrics interface
 */
export interface WorkbookMetrics {
  totalSheets: number;
  totalCells: number;
  totalFormulas: number;
  totalTables: number;
  totalCharts: number;
}

/**
 * Metadata chunk for efficient, selective capture and compression
 */
export interface MetadataChunk {
  id: string;            // e.g. "Sheet:Income Statement" or "Range:Sheet1!D2:D100"
  type: 'sheet' | 'range';
  etag: string;          // Hash of raw values+formulas+format for change detection
  payload: any;          // Compressed JSON - schema depends on chunk type
  summary?: string;      // 1-sentence summary for quick LLM reference (optional)
  refs: string[];        // Other ChunkIds this chunk depends on (for dependency graph)
  lastCaptured: Date;    // When this chunk was last captured/refreshed
}

/**
 * Query context built from relevant chunks for a specific query
 */
export interface QueryContext {
  chunks: MetadataChunk[];   // Only the chunks needed for this query
  activeSheet: string;       // Retained from existing model
  userSelection?: string;    // If user selected a range
  metrics: WorkbookMetrics;  // Aggregated over included chunks
}
