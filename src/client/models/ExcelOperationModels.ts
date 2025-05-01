// src/client/models/ExcelOperationModels.ts
// Defines the Excel Operations DSL for command execution

/**
 * Supported Excel operation types
 */
export enum ExcelOperationType {
  SET_VALUE = 'set_value',
  ADD_FORMULA = 'add_formula',
  CREATE_CHART = 'create_chart',
  FORMAT_RANGE = 'format_range',
  CLEAR_RANGE = 'clear_range',
  CREATE_TABLE = 'create_table',
  SORT_RANGE = 'sort_range',
  FILTER_RANGE = 'filter_range',
  CREATE_SHEET = 'create_sheet',
  DELETE_SHEET = 'delete_sheet',
  RENAME_SHEET = 'rename_sheet',
  COPY_RANGE = 'copy_range',
  MERGE_CELLS = 'merge_cells',
  UNMERGE_CELLS = 'unmerge_cells',
  CONDITIONAL_FORMAT = 'conditional_format',
  ADD_COMMENT = 'add_comment'
}

/**
 * Base interface for all Excel operations
 */
export interface BaseExcelOperation {
  op: ExcelOperationType;
}

/**
 * Set a value in a cell
 */
export interface SetValueOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_VALUE;
  target: string; // e.g. "Sheet1!A1"
  value: string | number | boolean;
}

/**
 * Add a formula to a cell
 */
export interface AddFormulaOperation extends BaseExcelOperation {
  op: ExcelOperationType.ADD_FORMULA;
  target: string; // e.g. "Sheet1!A1"
  formula: string; // e.g. "=SUM(B1:B10)"
}

/**
 * Create a chart
 */
export interface CreateChartOperation extends BaseExcelOperation {
  op: ExcelOperationType.CREATE_CHART;
  range: string; // e.g. "Sheet1!A1:D10"
  type: string; // e.g. "columnClustered", "line", "pie"
  title?: string;
  position?: string; // e.g. "Sheet1!F1"
}

/**
 * Format a range of cells
 */
export interface FormatRangeOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_RANGE;
  range: string; // e.g. "Sheet1!A1:D10"
  style?: string; // e.g. "Currency", "Percentage"
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
  fontSize?: number;
  horizontalAlignment?: string; // e.g. "left", "center", "right"
  verticalAlignment?: string; // e.g. "top", "center", "bottom"
}

/**
 * Clear a range of cells
 */
export interface ClearRangeOperation extends BaseExcelOperation {
  op: ExcelOperationType.CLEAR_RANGE;
  range: string; // e.g. "Sheet1!A1:D10"
  clearType?: string; // e.g. "all", "formats", "contents"
}

/**
 * Create a table
 */
export interface CreateTableOperation extends BaseExcelOperation {
  op: ExcelOperationType.CREATE_TABLE;
  range: string; // e.g. "Sheet1!A1:D10"
  hasHeaders?: boolean;
  styleName?: string; // e.g. "TableStyleMedium2"
}

/**
 * Sort a range
 */
export interface SortRangeOperation extends BaseExcelOperation {
  op: ExcelOperationType.SORT_RANGE;
  range: string; // e.g. "Sheet1!A1:D10"
  sortBy: string; // e.g. "A", "B", "C"
  sortDirection: string; // e.g. "ascending", "descending"
  hasHeaders?: boolean;
}

/**
 * Filter a range
 */
export interface FilterRangeOperation extends BaseExcelOperation {
  op: ExcelOperationType.FILTER_RANGE;
  range: string; // e.g. "Sheet1!A1:D10"
  column: string; // e.g. "A", "B", "C"
  criteria: string; // e.g. ">0", "=Red", "<>0"
}

/**
 * Create a new worksheet
 */
export interface CreateSheetOperation extends BaseExcelOperation {
  op: ExcelOperationType.CREATE_SHEET;
  name: string;
  position?: number; // 0-based index
}

/**
 * Delete a worksheet
 */
export interface DeleteSheetOperation extends BaseExcelOperation {
  op: ExcelOperationType.DELETE_SHEET;
  name: string;
}

/**
 * Rename a worksheet
 */
export interface RenameSheetOperation extends BaseExcelOperation {
  op: ExcelOperationType.RENAME_SHEET;
  oldName: string;
  newName: string;
}

/**
 * Copy a range to another location
 */
export interface CopyRangeOperation extends BaseExcelOperation {
  op: ExcelOperationType.COPY_RANGE;
  source: string; // e.g. "Sheet1!A1:D10"
  destination: string; // e.g. "Sheet2!A1"
}

/**
 * Merge cells
 */
export interface MergeCellsOperation extends BaseExcelOperation {
  op: ExcelOperationType.MERGE_CELLS;
  range: string; // e.g. "Sheet1!A1:D1"
}

/**
 * Unmerge cells
 */
export interface UnmergeCellsOperation extends BaseExcelOperation {
  op: ExcelOperationType.UNMERGE_CELLS;
  range: string; // e.g. "Sheet1!A1:D1"
}

/**
 * Add conditional formatting
 */
export interface ConditionalFormatOperation extends BaseExcelOperation {
  op: ExcelOperationType.CONDITIONAL_FORMAT;
  range: string; // e.g. "Sheet1!A1:D10"
  type: string; // e.g. "dataBar", "colorScale", "iconSet", "topBottom", "custom"
  criteria?: string;
  format?: {
    fontColor?: string;
    fillColor?: string;
    bold?: boolean;
    italic?: boolean;
  };
}

/**
 * Add a comment to a cell
 */
export interface AddCommentOperation extends BaseExcelOperation {
  op: ExcelOperationType.ADD_COMMENT;
  target: string; // e.g. "Sheet1!A1"
  text: string;
}

/**
 * Union type of all Excel operations
 */
export type ExcelOperation =
  | SetValueOperation
  | AddFormulaOperation
  | CreateChartOperation
  | FormatRangeOperation
  | ClearRangeOperation
  | CreateTableOperation
  | SortRangeOperation
  | FilterRangeOperation
  | CreateSheetOperation
  | DeleteSheetOperation
  | RenameSheetOperation
  | CopyRangeOperation
  | MergeCellsOperation
  | UnmergeCellsOperation
  | ConditionalFormatOperation
  | AddCommentOperation;

/**
 * Command plan with operations
 */
export interface ExcelCommandPlan {
  description: string;
  operations: ExcelOperation[];
}
