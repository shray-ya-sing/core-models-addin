// src/client/models/ExcelOperationModels.ts
// Defines the Excel Operations DSL for command execution

/**
 * Supported Excel operation types
 */
export enum ExcelOperationType {
  // Original operations
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
  ADD_COMMENT = 'add_comment',
  
  // Worksheet settings
  SET_GRIDLINES = 'set_gridlines',
  SET_HEADERS = 'set_headers',
  SET_ZOOM = 'set_zoom',
  SET_FREEZE_PANES = 'set_freeze_panes',
  SET_VISIBLE = 'set_visible',
  SET_ACTIVE_SHEET = 'set_active_sheet',
  
  // Print settings
  SET_PRINT_AREA = 'set_print_area',
  SET_PRINT_ORIENTATION = 'set_print_orientation',
  SET_PRINT_MARGINS = 'set_print_margins',
  SET_PRINT_SCALING = 'set_print_scaling',
  SET_PRINT_HEADERS = 'set_print_headers',
  SET_PRINT_TITLE_ROWS = 'set_print_title_rows',
  SET_PRINT_TITLE_COLUMNS = 'set_print_title_columns',
  
  // Chart formatting
  FORMAT_CHART = 'format_chart',
  FORMAT_CHART_AXIS = 'format_chart_axis',
  FORMAT_CHART_SERIES = 'format_chart_series',
  FORMAT_CHART_LEGEND = 'format_chart_legend',
  FORMAT_CHART_TITLE = 'format_chart_title',
  FORMAT_CHART_DATAPOINT = 'format_chart_datapoint',
  
  // Data operations
  ADD_DATA_VALIDATION = 'add_data_validation',
  ADD_SLICER = 'add_slicer',
  ADD_SPARKLINE = 'add_sparkline',
  
  // Composite and complex operations
  COMPOSITE_OPERATION = 'composite_operation',
  EXECUTE_SCRIPT = 'execute_script',
  BATCH_OPERATION = 'batch_operation'
}

/**
 * Base interface for all Excel operations
 */
export interface BaseExcelOperation {
  op: ExcelOperationType;
  id?: string;                // Optional identifier for the operation
  description?: string;       // Human-readable description of what this operation does
  dependsOn?: string[];      // IDs of operations this operation depends on
  ignoreErrors?: boolean;    // Whether to continue execution if this operation fails
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
 * Set worksheet gridlines visibility
 */
export interface SetGridlinesOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_GRIDLINES;
  sheet: string;             // Sheet name
  display: boolean;          // Whether to display gridlines
}

/**
 * Set worksheet headers visibility
 */
export interface SetHeadersOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_HEADERS;
  sheet: string;             // Sheet name
  display: boolean;          // Whether to display row/column headers
}

/**
 * Set worksheet zoom level
 */
export interface SetZoomOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_ZOOM;
  sheet: string;             // Sheet name
  zoomLevel: number;         // Zoom percentage (e.g., 100, 150, 75)
}

/**
 * Freeze panes in a worksheet
 */
export interface SetFreezePanesOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_FREEZE_PANES;
  sheet: string;             // Sheet name
  address?: string;          // Cell address to freeze at (e.g., "B3")
  row?: number;              // Row to freeze above (0 for none)
  column?: number;           // Column to freeze to the left (0 for none)
}

/**
 * Set worksheet visibility
 */
export interface SetVisibleOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_VISIBLE;
  sheet: string;             // Sheet name
  visible: boolean;          // Whether the sheet is visible
}

/**
 * Set the active worksheet
 */
export interface SetActiveSheetOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_ACTIVE_SHEET;
  sheet: string;             // Sheet name to activate
}

/**
 * Set print area for a worksheet
 */
export interface SetPrintAreaOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_AREA;
  sheet: string;             // Sheet name
  range: string;             // Range to set as print area (e.g., "A1:H20")
}

/**
 * Set print orientation for a worksheet
 */
export interface SetPrintOrientationOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_ORIENTATION;
  sheet: string;             // Sheet name
  orientation: 'portrait' | 'landscape';  // Print orientation
}

/**
 * Set print margins for a worksheet
 */
export interface SetPrintMarginsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_MARGINS;
  sheet: string;             // Sheet name
  top?: number;              // Top margin in inches
  right?: number;            // Right margin in inches
  bottom?: number;           // Bottom margin in inches
  left?: number;             // Left margin in inches
  header?: number;           // Header margin in inches
  footer?: number;           // Footer margin in inches
}

/**
 * Set print scaling for a worksheet
 */
export interface SetPrintScalingOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_SCALING;
  sheet: string;             // Sheet name
  scale?: number;            // Scale percentage (e.g., 100, 90, 75)
  fitToWidth?: number;       // Number of pages to fit width to
  fitToHeight?: number;      // Number of pages to fit height to
}

/**
 * Set print headers for a worksheet
 */
export interface SetPrintHeadersOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_HEADERS;
  sheet: string;             // Sheet name
  display: boolean;          // Whether to print row/column headers
}

/**
 * Set print title rows for a worksheet
 */
export interface SetPrintTitleRowsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_TITLE_ROWS;
  sheet: string;             // Sheet name
  range: string;             // Row range to repeat (e.g., "1:3")
}

/**
 * Set print title columns for a worksheet
 */
export interface SetPrintTitleColumnsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_TITLE_COLUMNS;
  sheet: string;             // Sheet name
  range: string;             // Column range to repeat (e.g., "A:C")
}

/**
 * Format a chart
 */
export interface FormatChartOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_CHART;
  sheet: string;             // Sheet name
  chartName?: string;        // Chart name (if known)
  chartIndex?: number;       // Chart index (if name unknown)
  title?: string;            // New chart title
  hasLegend?: boolean;       // Whether to show legend
  legendPosition?: 'top' | 'bottom' | 'left' | 'right' | 'corner';  // Legend position
  height?: number;           // Chart height in points
  width?: number;            // Chart width in points
  style?: number;            // Chart style (1-48)
  borderColor?: string;      // Chart border color
  borderWeight?: number;     // Chart border weight
}

/**
 * Format a chart axis
 */
export interface FormatChartAxisOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_CHART_AXIS;
  sheet: string;             // Sheet name
  chartName?: string;        // Chart name (if known)
  chartIndex?: number;       // Chart index (if name unknown)
  axisType: 'category' | 'value' | 'series';  // Axis type
  axisGroup?: 'primary' | 'secondary';  // Axis group
  title?: string;            // Axis title
  hasTitle?: boolean;        // Whether to show axis title
  showMajorGridlines?: boolean;  // Whether to show major gridlines
  showMinorGridlines?: boolean;  // Whether to show minor gridlines
  majorUnit?: number;        // Major unit for axis scaling
  minorUnit?: number;        // Minor unit for axis scaling
  minimum?: number;          // Minimum value for axis
  maximum?: number;          // Maximum value for axis
  displayUnit?: string;      // Display unit (e.g., "thousands", "millions")
  logScale?: boolean;        // Whether to use logarithmic scale
  reversed?: boolean;        // Whether axis is reversed
  tickLabelPosition?: string;  // Tick label position
  tickMarkType?: string;     // Tick mark type
}

/**
 * Format a chart series
 */
export interface FormatChartSeriesOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_CHART_SERIES;
  sheet: string;             // Sheet name
  chartName?: string;        // Chart name (if known)
  chartIndex?: number;       // Chart index (if name unknown)
  seriesName?: string;       // Series name to format
  seriesIndex?: number;      // Series index to format
  lineColor?: string;        // Line color
  lineWeight?: number;       // Line weight
  markerStyle?: string;      // Marker style
  markerSize?: number;       // Marker size
  markerColor?: string;      // Marker color
  fillColor?: string;        // Fill color
  transparency?: number;     // Transparency (0-100)
  plotOrder?: number;        // Plot order
  gapWidth?: number;         // Gap width (for column/bar charts)
  overlap?: number;          // Overlap (for column/bar charts)
  bubble3D?: boolean;        // Whether to use 3D bubbles
  bubbleSize?: number;       // Bubble size
  explosive?: number;        // Explosive factor (for pie/doughnut)
}

/**
 * Format a chart legend
 */
export interface FormatChartLegendOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_CHART_LEGEND;
  sheet: string;             // Sheet name
  chartName?: string;        // Chart name (if known)
  chartIndex?: number;       // Chart index (if name unknown)
  position?: 'top' | 'bottom' | 'left' | 'right' | 'corner';  // Legend position
  visible?: boolean;         // Whether legend is visible
  fontSize?: number;         // Font size
  fontColor?: string;        // Font color
  fontBold?: boolean;        // Whether font is bold
  fontItalic?: boolean;      // Whether font is italic
}

/**
 * Format a chart title
 */
export interface FormatChartTitleOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_CHART_TITLE;
  sheet: string;             // Sheet name
  chartName?: string;        // Chart name (if known)
  chartIndex?: number;       // Chart index (if name unknown)
  text?: string;             // Title text
  visible?: boolean;         // Whether title is visible
  fontSize?: number;         // Font size
  fontColor?: string;        // Font color
  fontBold?: boolean;        // Whether font is bold
  fontItalic?: boolean;      // Whether font is italic
  position?: 'top' | 'centered' | 'bottom';  // Title position
}

/**
 * Format chart data points
 */
export interface FormatChartDataPointOperation extends BaseExcelOperation {
  op: ExcelOperationType.FORMAT_CHART_DATAPOINT;
  sheet: string;             // Sheet name
  chartName?: string;        // Chart name (if known)
  chartIndex?: number;       // Chart index (if name unknown)
  seriesIndex: number;       // Series index
  pointIndex: number;        // Point index
  fillColor?: string;        // Fill color
  borderColor?: string;      // Border color
  borderWeight?: number;     // Border weight
  explosive?: number;        // Explosive factor (for pie/doughnut)
  marker?: {                 // Marker properties
    style?: string;          // Marker style
    size?: number;           // Marker size
    color?: string;          // Marker color
  };
}

/**
 * Add data validation to a range
 */
export interface AddDataValidationOperation extends BaseExcelOperation {
  op: ExcelOperationType.ADD_DATA_VALIDATION;
  sheet: string;             // Sheet name
  range: string;             // Range to validate
  type: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';  // Validation type
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';  // Comparison operator
  formula1?: string;         // First formula (or source range for lists)
  formula2?: string;         // Second formula (for between operations)
  showError?: boolean;       // Whether to show error
  errorTitle?: string;       // Error title
  errorMessage?: string;     // Error message
  errorStyle?: 'information' | 'warning' | 'stop';  // Error style
  showInput?: boolean;       // Whether to show input message
  inputTitle?: string;       // Input title
  inputMessage?: string;     // Input message
}

/**
 * Add a slicer to a table
 */
export interface AddSlicerOperation extends BaseExcelOperation {
  op: ExcelOperationType.ADD_SLICER;
  sheet: string;             // Sheet name
  tableName: string;         // Table name
  columnName: string;        // Column name to slice by
  destinationSheet?: string; // Destination sheet for the slicer
  position?: string;         // Position for the slicer
  height?: number;           // Slicer height
  width?: number;            // Slicer width
  style?: string;            // Slicer style
  caption?: string;          // Slicer caption
}

/**
 * Add sparklines to cells
 */
export interface AddSparklineOperation extends BaseExcelOperation {
  op: ExcelOperationType.ADD_SPARKLINE;
  sheet: string;             // Sheet name
  dataRange: string;         // Data range for sparklines
  locationRange: string;     // Location for sparklines
  type?: 'line' | 'column' | 'winLoss';  // Sparkline type
  lineColor?: string;        // Line color
  markerColor?: string;      // Marker color
  highPointColor?: string;   // High point color
  lowPointColor?: string;    // Low point color
  firstPointColor?: string;  // First point color
  lastPointColor?: string;   // Last point color
  negativePointColor?: string;  // Negative point color
  displayEmptyCellsAs?: 'gaps' | 'zero' | 'connect';  // How to display empty cells
}

/**
 * Composite operation containing multiple sub-operations
 */
export interface CompositeOperation extends BaseExcelOperation {
  op: ExcelOperationType.COMPOSITE_OPERATION;
  name: string;              // Name of the composite operation
  subOperations: ExcelOperation[];  // List of operations to execute
  abortOnFailure?: boolean;  // Whether to abort if a sub-operation fails
}

/**
 * Execute custom script
 */
export interface ExecuteScriptOperation extends BaseExcelOperation {
  op: ExcelOperationType.EXECUTE_SCRIPT;
  script: string;            // JavaScript code to execute
  context?: any;             // Context/parameters for the script
}

/**
 * Batch operation for improved performance
 */
export interface BatchOperation extends BaseExcelOperation {
  op: ExcelOperationType.BATCH_OPERATION;
  operations: ExcelOperation[];  // Operations to batch together
  requiresSync?: boolean;    // Whether a sync is required after these operations
}

/**
 * Union type of all Excel operations
 */
export type ExcelOperation =
  // Original operations
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
  | AddCommentOperation
  
  // Worksheet settings
  | SetGridlinesOperation
  | SetHeadersOperation
  | SetZoomOperation
  | SetFreezePanesOperation
  | SetVisibleOperation
  | SetActiveSheetOperation
  
  // Print settings
  | SetPrintAreaOperation
  | SetPrintOrientationOperation
  | SetPrintMarginsOperation
  | SetPrintScalingOperation
  | SetPrintHeadersOperation
  | SetPrintTitleRowsOperation
  | SetPrintTitleColumnsOperation
  
  // Chart formatting
  | FormatChartOperation
  | FormatChartAxisOperation
  | FormatChartSeriesOperation
  | FormatChartLegendOperation
  | FormatChartTitleOperation
  | FormatChartDataPointOperation
  
  // Data operations
  | AddDataValidationOperation
  | AddSlicerOperation
  | AddSparklineOperation
  
  // Composite operations
  | CompositeOperation
  | ExecuteScriptOperation
  | BatchOperation;

/**
 * Command plan with operations
 */
export interface ExcelCommandPlan {
  id?: string;                // Optional plan identifier
  description: string;        // Human-readable description of the plan
  operations: ExcelOperation[];  // List of operations to execute
  metadata?: {               // Optional metadata
    creator?: string;         // Who/what created this plan
    created?: Date;           // When the plan was created
    purpose?: string;         // Purpose of this plan
    tags?: string[];          // Tags for categorization
  };
  error?: {                  // Error information if plan failed
    message: string;          // Error message
    operationId?: string;     // ID of the operation that failed
    details?: any;            // Additional error details
  };
}
