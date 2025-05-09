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
  SET_WORKSHEET_SETTINGS = 'set_worksheet_settings',
  SET_FREEZE_PANES = 'set_freeze_panes',
  SET_ACTIVE_SHEET = 'set_active_sheet',
  SET_TAB_COLOR = 'set_tab_color',
  
  // Print settings
  SET_PRINT_SETTINGS = 'set_print_settings',
  SET_PAGE_SETUP = 'set_page_setup',
  
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
  CREATE_SCENARIO_TABLE = 'create_scenario_table',
  SET_ROW_COLUMN_OPTIONS = 'set_row_column_options',
  
  // Composite and complex operations
  COMPOSITE_OPERATION = 'composite_operation',
  EXECUTE_SCRIPT = 'execute_script',
  BATCH_OPERATION = 'batch_operation',
  
  // PDF export operations
  EXPORT_TO_PDF = 'export_to_pdf',
  
  // Calculation options
  SET_CALCULATION_OPTIONS = 'set_calculation_options',
  RECALCULATE_RANGES = 'recalculate_ranges'
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
 * Set the active worksheet
 */
export interface SetActiveSheetOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_ACTIVE_SHEET;
  sheet: string;             // Sheet name to activate
}

/**
 * Set worksheet settings
 */
export interface SetWorksheetSettingsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_WORKSHEET_SETTINGS;
  sheet: string;             // Sheet name
  visible?: boolean;          // Whether the sheet is visible
  tabColor?: string;         // Color value (hex code, named color, or theme color)
  name?: string;             // New sheet name
}

/**
 * Set print settings for a worksheet
 */
export interface SetPrintSettingsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PRINT_SETTINGS;
  sheet: string;
  blackAndWhite?: boolean;       // Whether to print in black and white
  draftMode?: boolean;           // Whether to print in draft mode
  firstPageNumber?: number;      // First page number
  headings?: boolean;            // Whether to display row/column headings when printing
  orientation?: string;          // "portrait" or "landscape"
  printAreas?: string[];         // Ranges to set as print areas (e.g. ["A1:H20", "A20:H40"])
  printComments?: string;        // "none", "at_end", "as_displayed"
  headerMargin?: number;         // Header margin in inches
  footerMargin?: number;         // Footer margin in inches
  leftMargin?: number;           // Left margin in inches
  rightMargin?: number;          // Right margin in inches
  topMargin?: number;            // Top margin in inches
  bottomMargin?: number;         // Bottom margin in inches
  printErrors?: string;          // "blank", "dash", "displayed", "na"
  headerRows?: number;           // Number of header rows
  footerRows?: number;           // Number of footer rows 
  printTitles?: string[];        // Ranges to set as print titles (e.g. ["A1:H1", "A1:H1"])
  printGridlines?: boolean;      // Whether to display gridlines when printing
  paperSize?: string;            // Paper size (e.g. "letter", "legal", "a4")
  centerHorizontally?: boolean;  // Center horizontally on page
  centerVertically?: boolean;    // Center vertically on page
  scale?: number;                // Scale percentage
  fitToWidth?: number;           // Number of pages wide to fit
  fitToHeight?: number;          // Number of pages tall to fit
  printOrder?: string;           // "over_then_down" or "down_then_over"
  leftHeader?: string;           // Left header text
  centerHeader?: string;         // Center header text
  rightHeader?: string;          // Right header text
  leftFooter?: string;           // Left footer text
  centerFooter?: string;         // Center footer text
  rightFooter?: string;          // Right footer text
}

/**
 * Set page setup for a worksheet
 */
export interface SetPageSetupOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_PAGE_SETUP;
  sheet: string;
  pageLayoutView?: string;       // "print" or "view"
  zoom?: number;                 // Zoom percentage
  gridlines?: boolean;           // Whether to display gridlines
  headers?: boolean;             // Whether to display row/column headers
  showFormulas?: boolean;        // Whether to display formulas instead of values
  showHeadings?: boolean;        // Whether to display row/column headings
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
  height?: number;           // Chart height in points
  width?: number;            // Chart width in points
  left?: number;             // Chart left position
  top?: number;              // Chart top position
  style?: number;            // Chart style (1-48)
  // Fill color
  fillColor?: string;        // Chart fill color
  hasFill?: boolean;         // Whether chart has fill
  // Border properties
  hasBorder?: boolean;       // Whether chart has border
  borderColor?: string;      // Chart border color
  borderWeight?: number;     // Chart border weight
  borderStyle?: string;      // Chart border style
  borderDashStyle?: string;  // Chart border dash style
  // Chart type
  chartType?: string;        // Chart type (e.g., "columnClustered", "line", "pie")
  chartSubType?: string;     // Chart sub-type (e.g., "stacked", "stacked100")
  chartGroup?: string;       // Chart group (e.g., "primary", "secondary")
  
  showMajorGridlines?: boolean;  // Whether to show major gridlines
  showMinorGridlines?: boolean;  // Whether to show minor gridlines
  // Axis properties
  majorUnit?: number;        // Major unit for axis scaling
  minorUnit?: number;        // Minor unit for axis scaling
  minimum?: number;          // Minimum value for axis
  maximum?: number;          // Maximum value for axis
  displayUnit?: string;      // Display unit (e.g., "thousands", "millions")
  logScale?: boolean;        // Whether to use logarithmic scale
  reversed?: boolean;        // Whether axis is reversed
  tickLabelPosition?: string;  // Tick label position
  tickMarkType?: string;     // Tick mark type
  axisType: 'category' | 'value' | 'series';  // Axis type
  axisGroup?: 'primary' | 'secondary';  // Axis group
  // Chart series properties
  seriesName?: string;       // Series name to format
  seriesIndex?: number;      // Series index to format
  lineColor?: string;        // Line color
  lineWeight?: number;       // Line weight
  markerStyle?: string;      // Marker style
  markerSize?: number;       // Marker size
  markerColor?: string;      // Marker color
  seriesFillColor?: string;        // Fill color
  transparency?: number;     // Transparency (0-100)
  plotOrder?: number;        // Plot order
  gapWidth?: number;         // Gap width (for column/bar charts)
  gapDepth?: number;         // Gap depth (for 3D charts)
  // Chart title properties
  hasTitle?: boolean;        // Whether to show chart title
  titleColor?: string;       // Chart title color
  titleFontName?: string;    // Chart title font name
  titleFontSize?: number;    // Chart title font size
  titleFontStyle?: string;   // Chart title font style
  titleFontBold?: boolean;   // Chart title font bold
  titleFontItalic?: boolean; // Chart title font italic
  titlePosition?: 'top' | 'centered' | 'bottom';  // Title position
  // Chart legend properties
  legendColor?: string;      // Chart legend color
  legendFontName?: string;   // Chart legend font name
  legendFontSize?: number;   // Chart legend font size
  legendFontStyle?: string;  // Chart legend font style
  legendFontBold?: boolean;  // Chart legend font bold
  legendFontItalic?: boolean; // Chart legend font italic
  legendPosition?: 'top' | 'bottom' | 'left' | 'right' | 'corner';  // Legend position
  legendVisible?: boolean;         // Whether legend is visible
  // Chart data point properties
  daatPointfillColor?: string;        // Fill color
  dataPointborderColor?: string;      // Border color
  dataPointborderWeight?: number;     // Border weight
  explosive?: number;        // Explosive factor (for pie/doughnut)
  marker?: {                 // Marker properties
    style?: string;          // Marker style
    size?: number;           // Marker size
    color?: string;          // Marker color
  };
  dataPointVisible?: boolean;         // Whether data point is visible

  // Data label properties
  dataLabelVisible?: boolean;         // Whether data label is visible
  dataLabelPosition?: 'top' | 'bottom' | 'left' | 'right' | 'center';  // Data label position
  dataLabelFontName?: string;   // Data label font name
  dataLabelFontSize?: number;   // Data label font size
  dataLabelFontStyle?: string;  // Data label font style
  dataLabelFontBold?: boolean;  // Data label font bold
  dataLabelFontItalic?: boolean; // Data label font italic
  dataLabelFontColor?: string;  // Data label font color
  dataLabelFormat?: string;     // Data label format
  dataLabelSeparator?: string;  // Data label separator
  dataLabelNumberFormat?: string;  // Data label number format
  dataLabelNumberFormatLinkedToSource?: boolean;  // Whether number format is linked to source
  dataLabelLeft?: number;        // Data label left position
  dataLabelTop?: number;         // Data label top position
  dataLabelWidth?: number;       // Data label width
  dataLabelHeight?: number;      // Data label height
  dataLabelShowCategoryName?: boolean;  // Whether to show category name
  dataLabelShowSeriesName?: boolean;  // Whether to show series name
  dataLabelShowValue?: boolean;  // Whether to show value
  dataLabelShowPercentage?: boolean;  // Whether to show percentage
  dataLabelShowBubbleSize?: boolean;  // Whether to show bubble size
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
  operations: ExcelOperation[];
  requiresSync?: boolean;
}

/**
 * Export worksheet to PDF
 */
export interface ExportToPdfOperation extends BaseExcelOperation {
  op: ExcelOperationType.EXPORT_TO_PDF;
  sheet: string;         // Sheet name to export
  fileName?: string;     // Name of the PDF file (without extension)
  filePath?: string;     // Optional file path where to save the PDF
  quality?: string;      // PDF quality: 'standard' or 'minimal'
  includeComments?: boolean; // Whether to include comments
  printArea?: string;    // Optional print area to export
  orientation?: 'portrait' | 'landscape'; // Page orientation
  fitToPage?: boolean;   // Whether to fit content to page
  margins?: {           // Optional page margins in points
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };
}

/**
 * Create a scenario table with sticky-IF formulas
 */
export interface CreateScenarioTableOperation extends BaseExcelOperation {
  op: ExcelOperationType.CREATE_SCENARIO_TABLE;
  range: string;              // Range where to create the scenario table
  formulaCell: string;        // Cell with the formula to evaluate
  inputCell: string;          // Cell that will be modified for each scenario
  values: (string | number)[]; // Values to use for scenarios
  tableType?: 'one-way' | 'two-way'; // Type of scenario table (default: 'one-way')
  rowInputCell?: string;      // Second input cell for two-way tables
  rowValues?: (string | number)[]; // Values for the second input (row input in two-way tables)
  includeFormula?: boolean;   // Whether to include the formula text in output (default: false)
  format?: {                  // Optional formatting options
    headerFormatting?: boolean; // Format headers with bold, etc.
    resultFormatting?: string;  // Format for result cells (e.g., 'currency', 'percentage')
    tableStyle?: string;        // Table style to apply
  };
  runScenarios?: boolean;     // Whether to immediately run through all scenarios (default: false)
}

/**
 * Operation to set row or column options including size, grouping, and visibility
 */
export interface SetRowColumnOptionsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_ROW_COLUMN_OPTIONS;
  sheet: string;                // Worksheet name
  type: 'row' | 'column';       // Whether operation applies to rows or columns
  indices: number[];            // Zero-based indices of rows/columns to modify
  
  // Size options
  size?: number;                // Width (for columns) or height (for rows) in points
  autofit?: boolean;            // Whether to autofit the rows/columns to their content
  
  // Grouping options
  group?: {
    start: number;              // Start index of the group (inclusive)
    end: number;                // End index of the group (inclusive)
    collapsed?: boolean;        // Whether to collapse the group (default: false)
  }[];
  
  // Visibility options
  hidden?: boolean;             // Whether to hide the rows/columns
  
  // Expanded state (for existing groups)
  expand?: boolean;             // Whether to expand existing groups that include these indices
}

/**
 * Operation to set calculation options
 */
export interface SetCalculationOptionsOperation extends BaseExcelOperation {
  op: ExcelOperationType.SET_CALCULATION_OPTIONS;
  calculationMode?: Excel.CalculationMode;
  iterative?: boolean;
  maxIterations?: number;
  maxChange?: number;
  calculate?: boolean;
  calculationType?: Excel.CalculationType;
}

/**
 * Operation to recalculate ranges
 */
export interface RecalculateRangesOperation extends BaseExcelOperation {
  op: ExcelOperationType.RECALCULATE_RANGES;
  recalculateAll?: boolean;
  sheets?: string[];
  ranges?: string[];
  cell?: string;
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
  | CopyRangeOperation
  | MergeCellsOperation
  | UnmergeCellsOperation
  | ConditionalFormatOperation
  | AddCommentOperation
  | SetFreezePanesOperation
  | SetActiveSheetOperation
  | SetPrintSettingsOperation
  | SetPageSetupOperation
  | FormatChartOperation
  | AddDataValidationOperation
  | AddSlicerOperation
  | AddSparklineOperation
  | CompositeOperation
  | ExecuteScriptOperation
  | BatchOperation
  | ExportToPdfOperation
  | SetWorksheetSettingsOperation
  | CreateScenarioTableOperation
  | SetRowColumnOptionsOperation
  | SetCalculationOptionsOperation
  | RecalculateRangesOperation;

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
