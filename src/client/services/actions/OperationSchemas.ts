import { z } from "zod";

// Common schema components
const sheetNameSchema = z.string();
const rangeSchema = z.string();
// Fix these shared schema components
const booleanOptional = z.boolean().optional().nullable();
const stringOptional = z.string().optional().nullable();
const numberOptional = z.number().optional().nullable();

// Base operation schema with discriminated union pattern
const baseOperationSchema = z.object({
  op: z.enum([
    'set_value', 'add_formula', 'create_chart', 'format_range', 'clear_range', 
    'create_table', 'sort_range', 'filter_range', 'create_sheet', 'delete_sheet',
    'copy_range', 'merge_cells', 'unmerge_cells', 'conditional_format', 'add_comment',
    'set_freeze_panes', 'set_active_sheet', 'set_print_settings', 'set_page_setup',
    'export_to_pdf', 'set_worksheet_settings', 'format_chart', 'set_calculation_options',
    'recalculate_ranges'
  ])
});

// 1. Set Value Operation
const setValueOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_value'),
  target: z.string(), // Cell reference (e.g. "Sheet1!A1")
  value: z.union([z.string(), z.number(), z.boolean()]) // Value to set (string, number, boolean)
});

// 2. Add Formula Operation
const addFormulaOperationSchema = baseOperationSchema.extend({
  op: z.literal('add_formula'),
  target: z.string(), // Cell reference (e.g. "Sheet1!A1")
  formula: z.string() // Formula to add (e.g. "=SUM(B1:B10)")
});

// 3. Create Chart Operation
const createChartOperationSchema = baseOperationSchema.extend({
  op: z.literal('create_chart'),
  range: rangeSchema, // Range for chart data (e.g. "Sheet1!A1:D10")
  type: z.string(), // Chart type (e.g. "columnClustered", "line", "pie")
  title: stringOptional, // Optional chart title
  position: stringOptional // Optional position (e.g. "Sheet1!F1")
});

// 4. Format Range Operation
const formatRangeOperationSchema = baseOperationSchema.extend({
  op: z.literal('format_range'),
  range: rangeSchema, // Range to format (e.g. "Sheet1!A1:D10")
  style: stringOptional, // Optional number format (e.g. "Currency", "Percentage")
  bold: booleanOptional, // Optional bold formatting
  italic: booleanOptional, // Optional italic formatting
  fontColor: stringOptional, // Optional font color
  fillColor: stringOptional, // Optional fill color
  fontSize: numberOptional, // Optional font size
  horizontalAlignment: z.enum(['left', 'center', 'right']).optional().nullable(), // Optional alignment
  verticalAlignment: z.enum(['top', 'center', 'bottom']).optional().nullable() // Optional alignment
});

// 5. Clear Range Operation
const clearRangeOperationSchema = baseOperationSchema.extend({
  op: z.literal('clear_range'),
  range: rangeSchema, // Range to clear (e.g. "Sheet1!A1:D10")
  clearType: z.enum(['all', 'formats', 'contents']).optional().nullable() // Optional clear type
});

// 6. Create Table Operation
const createTableOperationSchema = baseOperationSchema.extend({
  op: z.literal('create_table'),
  range: rangeSchema, // Range for table (e.g. "Sheet1!A1:D10")
  hasHeaders: booleanOptional, // Optional whether first row contains headers
  styleName: stringOptional // Optional table style name
});

// 7. Sort Range Operation
const sortRangeOperationSchema = baseOperationSchema.extend({
  op: z.literal('sort_range'),
  range: rangeSchema, // Range to sort (e.g. "Sheet1!A1:D10")
  sortBy: z.string(), // Column to sort by (e.g. "A", "B", "C")
  sortDirection: z.enum(['ascending', 'descending']), // Sort direction
  hasHeaders: booleanOptional // Optional whether first row contains headers
});

// 8. Filter Range Operation
const filterRangeOperationSchema = baseOperationSchema.extend({
  op: z.literal('filter_range'),
  range: rangeSchema, // Range to filter (e.g. "Sheet1!A1:D10")
  column: z.string(), // Column to filter (e.g. "A", "B", "C")
  criteria: z.string() // Filter criteria (e.g. ">0", "=Red", "<>0")
});

// 9. Create Sheet Operation
const createSheetOperationSchema = baseOperationSchema.extend({
  op: z.literal('create_sheet'),
  name: z.string(), // Name for the new sheet
  position: numberOptional // Optional position (0-based index)
});

// 10. Delete Sheet Operation
const deleteSheetOperationSchema = baseOperationSchema.extend({
  op: z.literal('delete_sheet'),
  name: z.string() // Name of the sheet to delete
});

// 11. Copy Range Operation
const copyRangeOperationSchema = baseOperationSchema.extend({
  op: z.literal('copy_range'),
  source: rangeSchema, // Source range (e.g. "Sheet1!A1:D10")
  destination: z.string() // Destination cell (e.g. "Sheet2!A1")
});

// 12. Merge Cells Operation
const mergeCellsOperationSchema = baseOperationSchema.extend({
  op: z.literal('merge_cells'),
  range: rangeSchema // Range to merge (e.g. "Sheet1!A1:D1")
});

// 13. Unmerge Cells Operation
const unmergeCellsOperationSchema = baseOperationSchema.extend({
  op: z.literal('unmerge_cells'),
  range: rangeSchema // Range to unmerge (e.g. "Sheet1!A1:D1")
});

// 14. Conditional Format Operation
const conditionalFormatOperationSchema = baseOperationSchema.extend({
  op: z.literal('conditional_format'),
  range: rangeSchema, // Range to format (e.g. "Sheet1!A1:D10")
  type: z.enum(['dataBar', 'colorScale', 'iconSet', 'topBottom', 'custom']), // Format type
  criteria: stringOptional, // Optional criteria for custom formats
  format: z.object({
    fontColor: stringOptional,
    fillColor: stringOptional, 
    bold: booleanOptional,
    italic: booleanOptional
  }).optional().nullable() // Optional format settings for custom formats
});

// 15. Add Comment Operation
const addCommentOperationSchema = baseOperationSchema.extend({
  op: z.literal('add_comment'),
  target: z.string(), // Cell reference (e.g. "Sheet1!A1")
  text: z.string() // Comment text
});

// 16. Set Freeze Panes Operation
const setFreezePanesOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_freeze_panes'),
  sheet: sheetNameSchema, // Sheet name
  address: z.string(), // Cell address to freeze at (e.g. "B3")
  freeze: z.boolean() // Whether to freeze panes
});

// 17. Set Print Settings Operation
const setPrintSettingsOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_print_settings'),
  sheet: sheetNameSchema, // Sheet name
  blackAndWhite: booleanOptional, // Whether to print in black and white
  draftMode: booleanOptional, // Whether to print in draft mode
  firstPageNumber: numberOptional, // First page number
  headings: booleanOptional, // Whether to display row/column headings when printing
  orientation: z.enum(['portrait', 'landscape']).optional().nullable(), // Orientation
  printAreas: z.array(z.string()).optional().nullable(), // Ranges to set as print areas
  printComments: z.enum(['none', 'at_end', 'as_displayed']).optional().nullable(), // Print comments
  headerMargin: numberOptional, // Header margin in inches
  footerMargin: numberOptional, // Footer margin in inches
  leftMargin: numberOptional, // Left margin in inches
  rightMargin: numberOptional, // Right margin in inches
  topMargin: numberOptional, // Top margin in inches
  bottomMargin: numberOptional, // Bottom margin in inches
  printErrors: z.enum(['blank', 'dash', 'displayed', 'na']).optional().nullable(), // Print errors
  headerRows: numberOptional, // Number of header rows
  footerRows: numberOptional, // Number of footer rows
  printTitles: z.array(z.string()).optional().nullable(), // Ranges to set as print titles
  printGridlines: booleanOptional // Whether to display gridlines when printing
});

// 18. Set Page Setup Operation
const setPageSetupOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_page_setup'),
  sheet: sheetNameSchema, // Sheet name
  pageLayoutView: z.enum(['print', 'normal', 'pageBreakPreview']).optional().nullable(), // Page layout view
  zoom: numberOptional, // Zoom percentage
  gridlines: booleanOptional, // Whether to display gridlines
  headers: booleanOptional, // Whether to display row and column headers
  showFormulas: booleanOptional, // Whether to display formulas instead of values
  showHeadings: booleanOptional // Whether to display row and column headings
});

// 19. Export to PDF Operation
const exportToPdfOperationSchema = baseOperationSchema.extend({
  op: z.literal('export_to_pdf'),
  sheet: sheetNameSchema, // Sheet name to export
  fileName: stringOptional, // Optional name for the PDF file (without extension)   
  quality: z.enum(['standard', 'minimal']).optional().nullable(), // Optional PDF quality
  includeComments: booleanOptional, // Optional: whether to include comments
  printArea: stringOptional, // Optional print area to export (e.g., "A1:H20")
  orientation: z.enum(['portrait', 'landscape']).optional().nullable(), // Optional page orientation
  fitToPage: booleanOptional, // Optional: whether to fit content to page
  margins: z.object({
    top: numberOptional,
    right: numberOptional,
    bottom: numberOptional,
    left: numberOptional
  }).optional().nullable() // Optional page margins in points   
});

// 20. Set Worksheet Settings Operation
const setWorksheetSettingsOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_worksheet_settings'),
  sheet: sheetNameSchema, // Sheet name
  pageLayoutView: z.enum(['print', 'normal', 'pageBreakPreview']).optional().nullable(), // Page layout view
  zoom: numberOptional, // Zoom percentage
  gridlines: booleanOptional, // Whether to display gridlines
  headers: booleanOptional, // Whether to display row and column headers
  showFormulas: booleanOptional, // Whether to display formulas instead of values
  showHeadings: booleanOptional, // Whether to display row and column headings
  position: z.number().int().optional().nullable(), // index of the sheet in the whole workbook (0 based)
  enableCalculation: booleanOptional, // Whether to enable calculation
  visibility: booleanOptional // Whether to make the sheet visible
});

// 21. Format Chart Operation
const formatChartOperationSchema = baseOperationSchema.extend({
  op: z.literal('format_chart'),
  sheet: sheetNameSchema, // Sheet name
  chart: z.string(), // Chart name
  title: stringOptional, // Chart title
  type: stringOptional, // Chart type
  dataSource: stringOptional, // Chart data source cell range address
  legend: booleanOptional, // Chart legend
  axis: booleanOptional, // Chart axis
  series: stringOptional, // Chart series
  dataLabels: booleanOptional, // Chart data labels
  
  // Chart Dimension properties
  width: numberOptional, // Chart width
  height: numberOptional, // Chart height
  
  // Chart position properties
  left: numberOptional, // Chart left position
  top: numberOptional, // Chart top position

  // Chart format properties
  fillColor: stringOptional, // Chart fill color
  borderVisible: booleanOptional, // Chart border visibility
  borderColor: stringOptional, // Chart border color
  borderWidth: numberOptional, // Chart border width
  borderStyle: stringOptional, // Chart border style
  borderDashStyle: stringOptional, // Chart border dash style
  
  // Chart title properties
  titleVisible: booleanOptional, // Whether title is visible
  titleFontName: stringOptional, // Title font name
  titleFontSize: numberOptional, // Title font size
  titleFontStyle: stringOptional, // Title font style
  titleFontBold: booleanOptional, // Title font bold
  titleFontItalic: booleanOptional, // Title font italic
  titleFontColor: stringOptional, // Title font color
  titleFormat: stringOptional, // Title format

  // Legend properties
  legendVisible: booleanOptional, // Whether legend is visible
  legendFontName: stringOptional, // Legend font name
  legendFontSize: numberOptional, // Legend font size
  legendFontStyle: stringOptional, // Legend font style
  legendFontBold: booleanOptional, // Legend font bold
  legendFontItalic: booleanOptional, // Legend font italic
  legendFontColor: stringOptional, // Legend font color
  legendFormat: stringOptional, // Legend format

  // Chart axis properties
  axisVisible: booleanOptional, // Whether axis is visible
  axisFontName: stringOptional, // Axis font name
  axisFontSize: numberOptional, // Axis font size
  axisFontStyle: stringOptional, // Axis font style
  axisFontBold: booleanOptional, // Axis font bold
  axisFontItalic: booleanOptional, // Axis font italic
  axisFontColor: stringOptional, // Axis font color
  axisFormat: stringOptional, // Axis format

  // Chart series properties
  seriesVisible: booleanOptional, // Whether series is visible
  seriesFontName: stringOptional, // Series font name
  seriesFontSize: numberOptional, // Series font size
  seriesFontStyle: stringOptional, // Series font style
  seriesFontBold: booleanOptional, // Series font bold
  seriesFontItalic: booleanOptional, // Series font italic
  seriesFontColor: stringOptional, // Series font color
  seriesFormat: stringOptional, // Series format

  // Chart data labels properties
  dataLabelsVisible: booleanOptional, // Whether data labels are visible
  dataLabelsFontName: stringOptional, // Data labels font name
  dataLabelsFontSize: numberOptional, // Data labels font size
  dataLabelsFontStyle: stringOptional, // Data labels font style
  dataLabelsFontBold: booleanOptional, // Data labels font bold
  dataLabelsFontItalic: booleanOptional, // Data labels font italic
  dataLabelsFontColor: stringOptional, // Data labels font color
  dataLabelsFormat: stringOptional // Data labels format
});

// 22. Set Calculation Options Operation
const setCalculationOptionsOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_calculation_options'),
  calculationMode: z.enum(['auto', 'manual']).optional().nullable(), // Optional calculation mode
  iterative: booleanOptional, // Optional whether to enable iterative calculation
  maxIterations: numberOptional, // Optional maximum number of iterations
  maxChange: numberOptional, // Optional maximum change for iterative calculation
  calculate: booleanOptional, // Optional whether to calculate the workbook
  calculationType: z.enum(['full', 'full_recalculate', 'recalculate']).optional().nullable() // Optional calculation type
});

// 23. Recalculate Ranges Operation
const recalculateRangesOperationSchema = baseOperationSchema.extend({
  op: z.literal('recalculate_ranges'),
  recalculateAll: booleanOptional, // Optional whether to recalculate all sheets
  sheets: z.array(z.string()).optional().nullable(), // List of sheet names to recalculate
  ranges: z.array(z.string()).optional().nullable() // List of cell range addresses to recalculate
});

// Set Active Sheet Operation
const setActiveSheetOperationSchema = baseOperationSchema.extend({
  op: z.literal('set_active_sheet'),
  name: z.string() // Name of the sheet to activate
});

// Union of all operation types
const operationSchema = z.discriminatedUnion('op', [
  setValueOperationSchema,
  addFormulaOperationSchema,
  createChartOperationSchema,
  formatRangeOperationSchema,
  clearRangeOperationSchema,
  createTableOperationSchema,
  sortRangeOperationSchema,
  filterRangeOperationSchema,
  createSheetOperationSchema,
  deleteSheetOperationSchema,
  copyRangeOperationSchema,
  mergeCellsOperationSchema,
  unmergeCellsOperationSchema,
  conditionalFormatOperationSchema,
  addCommentOperationSchema,
  setFreezePanesOperationSchema,
  setActiveSheetOperationSchema,
  setPrintSettingsOperationSchema,
  setPageSetupOperationSchema,
  exportToPdfOperationSchema,
  setWorksheetSettingsOperationSchema,
  formatChartOperationSchema,
  setCalculationOptionsOperationSchema,
  recalculateRangesOperationSchema
]);

// Final schema for the entire Excel command plan
export const excelCommandPlanSchema = z.object({
  description: z.string(),
  operations: z.array(operationSchema)
});

export const excelCommandPlanSchemaJSON = {
  "name": "command_plan",
  "type": "json_schema",
  "strict": true,
  "schema": {
    "type": "object",
    "properties": {
      "description": { "type": "string" },
        "operations": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "op": { "type": "string" },
              "target": { "type": ["string", "null"] },
              "value": { "type": ["string", "number", "boolean", "null"] },
              "formula": { "type": ["string", "null"] },
              "range": { "type": ["string", "null"] },
              "type": { "type": ["string", "null"] },
              "title": { "type": ["string", "null"] },
              "position": { "type": ["string", "number", "null"] },
              "style": { "type": ["string", "null"] },
              "bold": { "type": ["boolean", "null"] },
              "italic": { "type": ["boolean", "null"] },
              "fontColor": { "type": ["string", "null"] },
              "fillColor": { "type": ["string", "null"] },
              "fontSize": { "type": ["number", "null"] },
              "horizontalAlignment": { "type": ["string", "null"] },
              "verticalAlignment": { "type": ["string", "null"] },
              "clearType": { "type": ["string", "null"] },
              "hasHeaders": { "type": ["boolean", "null"] },
              "styleName": { "type": ["string", "null"] },
              "sortBy": { "type": ["string", "null"] },
              "sortDirection": { "type": ["string", "null"] },
              "column": { "type": ["string", "null"] },
              "criteria": { "type": ["string", "null"] },
              "name": { "type": ["string", "null"] },
              "source": { "type": ["string", "null"] },
              "destination": { "type": ["string", "null"] },
              "text": { "type": ["string", "null"] },
              "sheet": { "type": ["string", "null"] },
              "address": { "type": ["string", "null"] },
              "freeze": { "type": ["boolean", "null"] },
              "blackAndWhite": { "type": ["boolean", "null"] },
              "draftMode": { "type": ["boolean", "null"] },
              "firstPageNumber": { "type": ["number", "null"] },
              "headings": { "type": ["boolean", "null"] },
              "orientation": { "type": ["string", "null"] },
              "printAreas": { "type": ["array", "null"], "items": { "type": "string" } },
              "printComments": { "type": ["string", "null"] },
              "headerMargin": { "type": ["number", "null"] },
              "footerMargin": { "type": ["number", "null"] },
              "leftMargin": { "type": ["number", "null"] },
              "rightMargin": { "type": ["number", "null"] },
              "topMargin": { "type": ["number", "null"] },
              "bottomMargin": { "type": ["number", "null"] },
              "printErrors": { "type": ["string", "null"] },
              "headerRows": { "type": ["number", "null"] },
              "footerRows": { "type": ["number", "null"] },
              "printTitles": { "type": ["array", "null"], "items": { "type": "string" } },
              "printGridlines": { "type": ["boolean", "null"] },
              "pageLayoutView": { "type": ["string", "null"] },
              "zoom": { "type": ["number", "null"] },
              "gridlines": { "type": ["boolean", "null"] },
              "headers": { "type": ["boolean", "null"] },
              "showFormulas": { "type": ["boolean", "null"] },
              "showHeadings": { "type": ["boolean", "null"] },
              "fileName": { "type": ["string", "null"] },
              "quality": { "type": ["string", "null"] },
              "includeComments": { "type": ["boolean", "null"] },
              "printArea": { "type": ["string", "null"] },
              "fitToPage": { "type": ["boolean", "null"] },
              "margins": { 
                "type": ["object", "null"],
                "properties": {
                  "top": { "type": ["number", "null"] },
                  "right": { "type": ["number", "null"] },
                  "bottom": { "type": ["number", "null"] },
                  "left": { "type": ["number", "null"] }
                },
                "additionalProperties": false
              },
              "enableCalculation": { "type": ["boolean", "null"] },
              "visibility": { "type": ["boolean", "null"] },
              "chart": { "type": ["string", "null"] },
              "dataSource": { "type": ["string", "null"] },
              "legend": { "type": ["boolean", "null"] },
              "axis": { "type": ["boolean", "null"] },
              "series": { "type": ["string", "null"] },
              "dataLabels": { "type": ["boolean", "null"] },
              "width": { "type": ["number", "null"] },
              "height": { "type": ["number", "null"] },
              "left": { "type": ["number", "null"] },
              "top": { "type": ["number", "null"] },
              "borderVisible": { "type": ["boolean", "null"] },
              "borderColor": { "type": ["string", "null"] },
              "borderWidth": { "type": ["number", "null"] },
              "borderStyle": { "type": ["string", "null"] },
              "borderDashStyle": { "type": ["string", "null"] },
              "titleVisible": { "type": ["boolean", "null"] },
              "titleFontName": { "type": ["string", "null"] },
              "titleFontSize": { "type": ["number", "null"] },
              "titleFontStyle": { "type": ["string", "null"] },
              "titleFontBold": { "type": ["boolean", "null"] },
              "titleFontItalic": { "type": ["boolean", "null"] },
              "titleFontColor": { "type": ["string", "null"] },
              "titleFormat": { "type": ["string", "null"] },
              "legendVisible": { "type": ["boolean", "null"] },
              "legendFontName": { "type": ["string", "null"] },
              "legendFontSize": { "type": ["number", "null"] },
              "legendFontStyle": { "type": ["string", "null"] },
              "legendFontBold": { "type": ["boolean", "null"] },
              "legendFontItalic": { "type": ["boolean", "null"] },
              "legendFontColor": { "type": ["string", "null"] },
              "legendFormat": { "type": ["string", "null"] },
              "axisVisible": { "type": ["boolean", "null"] },
              "axisFontName": { "type": ["string", "null"] },
              "axisFontSize": { "type": ["number", "null"] },
              "axisFontStyle": { "type": ["string", "null"] },
              "axisFontBold": { "type": ["boolean", "null"] },
              "axisFontItalic": { "type": ["boolean", "null"] },
              "axisFontColor": { "type": ["string", "null"] },
              "axisFormat": { "type": ["string", "null"] },
              "seriesVisible": { "type": ["boolean", "null"] },
              "seriesFontName": { "type": ["string", "null"] },
              "seriesFontSize": { "type": ["number", "null"] },
              "seriesFontStyle": { "type": ["string", "null"] },
              "seriesFontBold": { "type": ["boolean", "null"] },
              "seriesFontItalic": { "type": ["boolean", "null"] },
              "seriesFontColor": { "type": ["string", "null"] },
              "seriesFormat": { "type": ["string", "null"] },
              "dataLabelsVisible": { "type": ["boolean", "null"] },
              "dataLabelsFontName": { "type": ["string", "null"] },
              "dataLabelsFontSize": { "type": ["number", "null"] },
              "dataLabelsFontStyle": { "type": ["string", "null"] },
              "dataLabelsFontBold": { "type": ["boolean", "null"] },
              "dataLabelsFontItalic": { "type": ["boolean", "null"] },
              "dataLabelsFontColor": { "type": ["string", "null"] },
              "dataLabelsFormat": { "type": ["string", "null"] },
              "calculationMode": { "type": ["string", "null"] },
              "iterative": { "type": ["boolean", "null"] },
              "maxIterations": { "type": ["number", "null"] },
              "maxChange": { "type": ["number", "null"] },
              "calculate": { "type": ["boolean", "null"] },
              "calculationType": { "type": ["string", "null"] },
              "recalculateAll": { "type": ["boolean", "null"] },
              "sheets": { "type": ["array", "null"], "items": { "type": "string" } },
              "ranges": { "type": ["array", "null"], "items": { "type": "string" } }
            },
            "required": ["op"],
            "additionalProperties": false
          }
        }
      },
      "required": ["description", "operations"],
      "additionalProperties": false
    }
  }

  export const openAICompatibleCommandPlanSchema = {
    "type": "json_schema",
    "name": "command_plan",
    "strict": true,
    "schema": {
      "type": "object",
      "properties": {
        "description": { "type": "string" },
        "operations": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "op": { "type": "string" },
              "target": { "type": "string" },
              "value": { "type": ["string", "number", "boolean", "null"] },
              "formula": { "type": "string" },
              "name": { "type": "string" },
              "sheet": { "type": "string" },
              "range": { "type": "string" },
              "source": { "type": "string" },
              "destination": { "type": "string" }
            },
            "required": ["op"],
            "additionalProperties": true
          }
        }
      },
      "required": ["description", "operations"],
      "additionalProperties": false
    }
  }

  export const detailedOpenAICommandPlanSchema = {
    "type": "json_schema",
    "name": "command_plan",
    "strict": true,
    "schema": {
      "type": "object",
      "properties": {
        "description": { "type": "string" },
        "operations": {
          "type": "array",
          "items": {
            "oneOf": [
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_value"] },
                  "target": { "type": "string" },
                  "value": { "type": ["string", "number", "boolean"] }
                },
                "required": ["op", "target", "value"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["add_formula"] },
                  "target": { "type": "string" },
                  "formula": { "type": "string" }
                },
                "required": ["op", "target", "formula"],
                "additionalProperties": false
              }
            ]
          }
        }
      },
      "required": ["description", "operations"],
      "additionalProperties": false
    }
  }

  export const finalOpenAICommandPlanSchema = {
    "type": "json_schema",
    "name": "command_plan",
    "strict": true,
    "schema": {
      "type": "object",
      "properties": {
        "description": { "type": "string" },
        "operations": {
          "type": "array",
          "items": {
            "oneOf": [
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_value"] },
                  "target": { "type": "string" },
                  "value": { "type": ["string", "number", "boolean"] }
                },
                "required": ["op", "target", "value"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["add_formula"] },
                  "target": { "type": "string" },
                  "formula": { "type": "string" }
                },
                "required": ["op", "target", "formula"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["create_chart"] },
                  "range": { "type": "string" },
                  "type": { "type": "string" },
                  "title": { "type": "string" },
                  "position": { "type": "string" }
                },
                "required": ["op", "range", "type", "title", "position"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["format_range"] },
                  "range": { "type": "string" },
                  "style": { "type": "string" },
                  "bold": { "type": "boolean" },
                  "italic": { "type": "boolean" },
                  "fontColor": { "type": "string" },
                  "fillColor": { "type": "string" },
                  "fontSize": { "type": "number" },
                  "horizontalAlignment": { "type": "string" },
                  "verticalAlignment": { "type": "string" }
                },
                "required": ["op", "range", "style", "bold", "italic", "fontColor", "fillColor", "fontSize", "horizontalAlignment", "verticalAlignment"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["clear_range"] },
                  "range": { "type": "string" },
                  "clearType": { "type": "string" } 
                },
                "required": ["op", "range", "clearType"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["create_table"] },
                  "range": { "type": "string" },
                  "hasHeaders": { "type": "boolean" },
                  "styleName": { "type": "string" }
                },
                "required": ["op", "range", "hasHeaders", "styleName"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["sort_range"] },
                  "range": { "type": "string" },
                  "sortBy": { "type": "string" },
                  "sortDirection": { "type": "string" },
                  "hasHeaders": { "type": "boolean" }
                },
                "required": ["op", "range", "sortBy", "sortDirection", "hasHeaders"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["filter_range"] },
                  "range": { "type": "string" },
                  "column": { "type": "string" },
                  "criteria": { "type": "string" }
                },
                "required": ["op", "range", "column", "criteria"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["create_sheet"] },
                  "name": { "type": "string" },
                  "position": { "type": "number" }
                },
                "required": ["op", "name", "position"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["delete_sheet"] },
                  "name": { "type": "string" }
                },
                "required": ["op", "name"], 
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["copy_range"] },
                  "source": { "type": "string" },
                  "destination": { "type": "string" }
                },
                "required": ["op", "source", "destination"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["merge_cells"] },
                  "range": { "type": "string" }
                },
                "required": ["op", "range"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["unmerge_cells"] },
                  "range": { "type": "string" }
                },
                "required": ["op", "range"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["conditional_format"] },
                  "range": { "type": "string" },
                  "type": { "type": "string" },
                  "criteria": { "type": "string" },
                  "format": {
                    "type": "object",
                    "properties": {
                      "fontColor": { "type": "string" },
                      "fillColor": { "type": "string" },
                      "bold": { "type": "boolean" },
                      "italic": { "type": "boolean" }
                    },
                    "required": ["fontColor", "fillColor", "bold", "italic"],
                    "additionalProperties": false
                  }
                },
                "required": ["op", "range", "type", "criteria", "format"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["add_comment"] },
                  "target": { "type": "string" },
                  "text": { "type": "string" }
                },
                "required": ["op", "target", "text"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_freeze_panes"] },
                  "sheet": { "type": "string" },
                  "address": { "type": "string" },
                  "freeze": { "type": "boolean" }
                },
                "required": ["op", "sheet", "address", "freeze"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_print_settings"] },
                  "sheet": { "type": "string" },
                  "blackAndWhite": { "type": "boolean" },
                  "draftMode": { "type": "boolean" },
                  "firstPageNumber": { "type": "number" },
                  "headings": { "type": "boolean" },
                  "orientation": { "type": "string" },
                  "printAreas": { "type": "array", "items": { "type": "string" } },
                  "printComments": { "type": "string" },
                  "headerMargin": { "type": "number" },
                  "footerMargin": { "type": "number" },
                  "leftMargin": { "type": "number" },
                  "rightMargin": { "type": "number" },
                  "topMargin": { "type": "number" },
                  "bottomMargin": { "type": "number" },
                  "printErrors": { "type": "string" },
                  "headerRows": { "type": "number" },
                  "footerRows": { "type": "number" },
                  "printTitles": { "type": "array", "items": { "type": "string" } },
                  "printGridlines": { "type": "boolean" }
                },
                "required": ["op", "sheet", "blackAndWhite", "draftMode", "firstPageNumber", "headings", "orientation", "printAreas", "printComments", "headerMargin", "footerMargin", "leftMargin", "rightMargin", "topMargin", "bottomMargin", "printErrors", "headerRows", "footerRows", "printTitles", "printGridlines"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_page_setup"] },
                  "sheet": { "type": "string" },
                  "pageLayoutView": { "type": "string" },
                  "zoom": { "type": "number" },
                  "gridlines": { "type": "boolean" },
                  "headers": { "type": "boolean" },
                  "showFormulas": { "type": "boolean" },
                  "showHeadings": { "type": "boolean" }
                },
                "required": ["op", "sheet", "pageLayoutView", "zoom", "gridlines", "headers", "showFormulas", "showHeadings"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["export_to_pdf"] },
                  "sheet": { "type": "string" },
                  "fileName": { "type": "string" },
                  "quality": { "type": "string" },
                  "includeComments": { "type": "boolean" },
                  "printArea": { "type": "string" },
                  "orientation": { "type": "string" },
                  "fitToPage": { "type": "boolean" },
                  "margins": {
                    "type": "object",
                    "properties": {
                      "top": { "type": "number" },
                      "right": { "type": "number" },
                      "bottom": { "type": "number" },
                      "left": { "type": "number" }
                    },
                    "required": ["top", "right", "bottom", "left"],
                    "additionalProperties": false
                  }
                },
                "required": ["op", "sheet", "fileName", "quality", "includeComments", "printArea", "orientation", "fitToPage", "margins"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_worksheet_settings"] },
                  "sheet": { "type": "string" },
                  "pageLayoutView": { "type": "string" },
                  "zoom": { "type": "number" },
                  "gridlines": { "type": "boolean" },
                  "headers": { "type": "boolean" },
                  "showFormulas": { "type": "boolean" },
                  "showHeadings": { "type": "boolean" },
                  "position": { "type": "number" },
                  "enableCalculation": { "type": "boolean" },
                  "visibility": { "type": "boolean" }
                },
                "required": ["op", "sheet", "pageLayoutView", "zoom", "gridlines", "headers", "showFormulas", "showHeadings", "position", "enableCalculation", "visibility"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["format_chart"] },
                  "sheet": { "type": "string" },
                  "chart": { "type": "string" },
                  "title": { "type": "string" },
                  "type": { "type": "string" },
                  "dataSource": { "type": "string" },
                  "legend": { "type": "boolean" },
                  "axis": { "type": "boolean" },
                  "series": { "type": "string" },
                  "dataLabels": { "type": "boolean" },
                  "width": { "type": "number" },
                  "height": { "type": "number" },
                  "left": { "type": "number" },
                  "top": { "type": "number" },
                  "fillColor": { "type": "string" },
                  "borderVisible": { "type": "boolean" },
                  "borderColor": { "type": "string" },
                  "borderWidth": { "type": "number" },
                  "borderStyle": { "type": "string" },
                  "borderDashStyle": { "type": "string" },
                  "titleVisible": { "type": "boolean" },
                  "titleFontName": { "type": "string" },
                  "titleFontSize": { "type": "number" },
                  "titleFontStyle": { "type": "string" },
                  "titleFontBold": { "type": "boolean" },
                  "titleFontItalic": { "type": "boolean" },
                  "titleFontColor": { "type": "string" },
                  "titleFormat": { "type": "string" },
                  "legendVisible": { "type": "boolean" },
                  "legendFontName": { "type": "string" },
                  "legendFontSize": { "type": "number" },
                  "legendFontStyle": { "type": "string" },
                  "legendFontBold": { "type": "boolean" },
                  "legendFontItalic": { "type": "boolean" },
                  "legendFontColor": { "type": "string" },
                  "legendFormat": { "type": "string" },
                  "axisVisible": { "type": "boolean" },
                  "axisFontName": { "type": "string" },
                  "axisFontSize": { "type": "number" },
                  "axisFontStyle": { "type": "string" },
                  "axisFontBold": { "type": "boolean" },
                  "axisFontItalic": { "type": "boolean" },
                  "axisFontColor": { "type": "string" },
                  "axisFormat": { "type": "string" },
                  "seriesVisible": { "type": "boolean" },
                  "seriesFontName": { "type": "string" },
                  "seriesFontSize": { "type": "number" },
                  "seriesFontStyle": { "type": "string" },
                  "seriesFontBold": { "type": "boolean" },
                  "seriesFontItalic": { "type": "boolean" },
                  "seriesFontColor": { "type": "string" },
                  "seriesFormat": { "type": "string" },
                  "dataLabelsVisible": { "type": "boolean" },
                  "dataLabelsFontName": { "type": "string" },
                  "dataLabelsFontSize": { "type": "number" },
                  "dataLabelsFontStyle": { "type": "string" },
                  "dataLabelsFontBold": { "type": "boolean" },
                  "dataLabelsFontItalic": { "type": "boolean" },
                  "dataLabelsFontColor": { "type": "string" },
                  "dataLabelsFormat": { "type": "string" }
                },
                "required": ["op", "sheet", "chart", "title", "type", "dataSource", "legend", "axis", "series", "dataLabels", "width", "height", "left", "top", "fillColor", "borderVisible", "borderColor", "borderWidth", "borderStyle", "borderDashStyle", "titleVisible", "titleFontName", "titleFontSize", "titleFontStyle", "titleFontBold", "titleFontItalic", "titleFontColor", "titleFormat", "legendVisible", "legendFontName", "legendFontSize", "legendFontStyle", "legendFontBold", "legendFontItalic", "legendFontColor", "legendFormat", "axisVisible", "axisFontName", "axisFontSize", "axisFontStyle", "axisFontBold", "axisFontItalic", "axisFontColor", "axisFormat", "seriesVisible", "seriesFontName", "seriesFontSize", "seriesFontStyle", "seriesFontBold", "seriesFontItalic", "seriesFontColor", "seriesFormat", "dataLabelsVisible", "dataLabelsFontName", "dataLabelsFontSize", "dataLabelsFontStyle", "dataLabelsFontBold", "dataLabelsFontItalic", "dataLabelsFontColor", "dataLabelsFormat"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["set_calculation_options"] },
                  "calculationMode": { "type": "string" },
                  "iterative": { "type": "boolean" },
                  "maxIterations": { "type": "number" },
                  "maxChange": { "type": "number" },
                  "calculate": { "type": "boolean" },
                  "calculationType": { "type": "string" }
                },
                "required": ["op", "calculationMode", "iterative", "maxIterations", "maxChange", "calculate", "calculationType"],
                "additionalProperties": false
              },
              {
                "type": "object",
                "properties": {
                  "op": { "type": "string", "enum": ["recalculate_ranges"] },
                  "recalculateAll": { "type": "boolean" },
                  "sheets": { "type": "array", "items": { "type": "string" } },
                  "ranges": { "type": "array", "items": { "type": "string" } }
                },
                "required": ["op", "recalculateAll", "sheets", "ranges"],
                "additionalProperties": false
              }
            ]
          }
        }
      },
      "required": ["description", "operations"],
      "additionalProperties": false
    }
  }

  // Example of function definitions for Excel operations
export const excelOperationFunctions = [
  {
    name: "set_value",
    description: "Set a value in a cell or range",
    parameters: {
      type: "object",
      properties: {
        target: {
          type: "string",
          description: "The target cell or range in A1 notation"
        },
        value: {
          type: ["string", "number", "boolean"],
          description: "The value to set"
        }
      },
      required: ["target", "value"]
    }
  },
  {
    name: "add_formula",
    description: "Add a formula to a cell or range",
    parameters: {
      type: "object",
      properties: {
        target: {
          type: "string",
          description: "The target cell or range in A1 notation"
        },
        formula: {
          type: "string",
          description: "The Excel formula to add"
        }
      },
      required: ["target", "formula"]
    }
  }
]


export const excelOperationTools = [
  {
    type: "function",
    name: "set_value",
    description: "Set a value in a cell",
    parameters: {
      type: "object",
      properties: {
        target: { type: "string", description: "Cell reference (e.g. 'Sheet1!A1')" },
        value: { type: ["string", "number", "boolean"], description: "Value to set (string, number, boolean)" }
      },
      required: ["target", "value"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "add_formula",
    description: "Add a formula to a cell",
    parameters: {
      type: "object",
      properties: {
        target: { type: "string", description: "Cell reference (e.g. 'Sheet1!A1')" },
        formula: { type: "string", description: "Formula to add (e.g. '=SUM(B1:B10)')" }
      },
      required: ["target", "formula"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "create_chart",
    description: "Create a chart",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range for chart data (e.g. 'Sheet1!A1:D10')" },
        type: { type: "string", description: "Chart type (e.g. 'columnClustered', 'line', 'pie')" },
        title: { type: "string", description: "Chart title" },
        position: { type: "string", description: "Position (e.g. 'Sheet1!F1')" }
      },
      required: ["range", "type", "title", "position"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "format_range",
    description: "Format a range of cells",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to format (e.g. 'Sheet1!A1:D10')" },
        style: { type: "string", description: "Number format (e.g. 'Currency', 'Percentage')" },
        bold: { type: "boolean", description: "Bold formatting" },
        italic: { type: "boolean", description: "Italic formatting" },
        fontColor: { type: "string", description: "Font color" },
        fillColor: { type: "string", description: "Fill color" },
        fontSize: { type: "number", description: "Font size" },
        horizontalAlignment: { type: "string", description: "Horizontal alignment (e.g. 'left', 'center', 'right')" },
        verticalAlignment: { type: "string", description: "Vertical alignment (e.g. 'top', 'center', 'bottom')" }
      },
      required: ["range", "style", "bold", "italic", "fontColor", "fillColor", "fontSize", "horizontalAlignment", "verticalAlignment"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "clear_range",
    description: "Clear a range of cells",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to clear (e.g. 'Sheet1!A1:D10')" },
        clearType: { type: "string", description: "Clear type (e.g. 'all', 'formats', 'contents')" }
      },
      required: ["range", "clearType"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "create_table",
    description: "Create a table",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range for table (e.g. 'Sheet1!A1:D10')" },
        hasHeaders: { type: "boolean", description: "Whether first row contains headers" },
        styleName: { type: "string", description: "Table style name" }
      },
      required: ["range", "hasHeaders", "styleName"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "sort_range",
    description: "Sort a range",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to sort (e.g. 'Sheet1!A1:D10')" },
        sortBy: { type: "string", description: "Column to sort by (e.g. 'A', 'B', 'C')" },
        sortDirection: { type: "string", description: "Sort direction (e.g. 'ascending', 'descending')" },
        hasHeaders: { type: "boolean", description: "Whether first row contains headers" }
      },
      required: ["range", "sortBy", "sortDirection", "hasHeaders"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "filter_range",
    description: "Filter a range",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to filter (e.g. 'Sheet1!A1:D10')" },
        column: { type: "string", description: "Column to filter (e.g. 'A', 'B', 'C')" },
        criteria: { type: "string", description: "Filter criteria (e.g. '>0', '=Red', '<>0')" }
      },
      required: ["range", "column", "criteria"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "create_sheet",
    description: "Create a new worksheet",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", description: "Name for the new sheet" },
        position: { type: "number", description: "Position (0-based index)" }
      },
      required: ["name", "position"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "delete_sheet",
    description: "Delete a worksheet",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", description: "Name of the sheet to delete" }
      },
      required: ["name"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "copy_range",
    description: "Copy a range to another location",
    parameters: {
      type: "object",
      properties: {
        source: { type: "string", description: "Source range (e.g. 'Sheet1!A1:D10')" },
        destination: { type: "string", description: "Destination cell (e.g. 'Sheet2!A1')" }
      },
      required: ["source", "destination"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "merge_cells",
    description: "Merge cells",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to merge (e.g. 'Sheet1!A1:D1')" }
      },
      required: ["range"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "unmerge_cells",
    description: "Unmerge cells",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to unmerge (e.g. 'Sheet1!A1:D1')" }
      },
      required: ["range"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "conditional_format",
    description: "Add conditional formatting",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range to format (e.g. 'Sheet1!A1:D10')" },
        type: { type: "string", description: "Format type (e.g. 'dataBar', 'colorScale', 'iconSet', 'topBottom', 'custom')" },
        criteria: { type: "string", description: "Criteria for custom formats" },
        format: { 
          type: "object", 
          description: "Format settings for custom formats",
          properties: {
            fontColor: { type: "string" },
            fillColor: { type: "string" },
            bold: { type: "boolean" },
            italic: { type: "boolean" }
          },
          required: ["fontColor", "fillColor", "bold", "italic"],
          additionalProperties: false
        }
      },
      required: ["range", "type", "criteria", "format"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "add_comment",
    description: "Add a comment to a cell",
    parameters: {
      type: "object",
      properties: {
        target: { type: "string", description: "Cell reference (e.g. 'Sheet1!A1')" },
        text: { type: "string", description: "Comment text" }
      },
      required: ["target", "text"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "set_freeze_panes",
    description: "Freeze rows or columns",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        address: { type: "string", description: "Cell address to freeze at (e.g. 'B3')" },
        freeze: { type: "boolean", description: "Whether to freeze panes. True to freeze, false to unfreeze." }
      },
      required: ["sheet", "address", "freeze"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "set_active_sheet",
    description: "Set the active worksheet",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", description: "Name of the sheet to activate" }
      },
      required: ["name"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "set_print_settings",
    description: "Set print settings",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        blackAndWhite: { type: "boolean", description: "Whether to print in black and white" },
        draftMode: { type: "boolean", description: "Whether to print in draft mode" },
        firstPageNumber: { type: "number", description: "First page number" },
        headings: { type: "boolean", description: "Whether to display row/column headings when printing" },
        orientation: { type: "string", description: "'portrait' or 'landscape'" },
        printAreas: { type: "array", items: { type: "string" }, description: "Ranges to set as print areas (e.g. ['A1:H20', 'A20:H40'])" },
        printComments: { type: "string", description: "'none', 'at_end', 'as_displayed'" },
        headerMargin: { type: "number", description: "Header margin in inches" },
        footerMargin: { type: "number", description: "Footer margin in inches" },
        leftMargin: { type: "number", description: "Left margin in inches" },
        rightMargin: { type: "number", description: "Right margin in inches" },
        topMargin: { type: "number", description: "Top margin in inches" },
        bottomMargin: { type: "number", description: "Bottom margin in inches" },
        printErrors: { type: "string", description: "'blank', 'dash', 'displayed', 'na'" },
        headerRows: { type: "number", description: "Number of header rows" },
        footerRows: { type: "number", description: "Number of footer rows" },
        printTitles: { type: "array", items: { type: "string" }, description: "Ranges to set as print titles (e.g. ['A1:H1', 'A1:H1'])" },
        printGridlines: { type: "boolean", description: "Whether to display gridlines when printing" }
      },
      required: ["sheet", "blackAndWhite", "draftMode", "firstPageNumber", "headings", "orientation", "printAreas", "printComments", "headerMargin", "footerMargin", "leftMargin", "rightMargin", "topMargin", "bottomMargin", "printErrors", "headerRows", "footerRows", "printTitles", "printGridlines"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "set_page_setup",
    description: "Set page setup",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        pageLayoutView: { type: "string", description: "'print', 'normal', 'pageBreakPreview'" },
        zoom: { type: "number", description: "Zoom percentage" },
        gridlines: { type: "boolean", description: "Whether to display gridlines" },
        headers: { type: "boolean", description: "Whether to display row and column headers" },
        showFormulas: { type: "boolean", description: "Whether to display formulas instead of values" },
        showHeadings: { type: "boolean", description: "Whether to display row and column headings" }
      },
      required: ["sheet", "pageLayoutView", "zoom", "gridlines", "headers", "showFormulas", "showHeadings"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "export_to_pdf",
    description: "Export worksheet to PDF",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name to export" },
        fileName: { type: "string", description: "Name for the PDF file (without extension)" },
        quality: { type: "string", description: "PDF quality: 'standard' or 'minimal'" },
        includeComments: { type: "boolean", description: "Whether to include comments" },
        printArea: { type: "string", description: "Print area to export (e.g., 'A1:H20')" },
        orientation: { type: "string", description: "Page orientation: 'portrait' or 'landscape'" },
        fitToPage: { type: "boolean", description: "Whether to fit content to page" },
        margins: {
          type: "object",
          description: "Page margins in points",
          properties: {
            top: { type: "number" },
            right: { type: "number" },
            bottom: { type: "number" },
            left: { type: "number" }
          },
          required: ["top", "right", "bottom", "left"],
          additionalProperties: false
        }
      },
      required: ["sheet", "fileName", "quality", "includeComments", "printArea", "orientation", "fitToPage", "margins"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "set_worksheet_settings",
    description: "Set worksheet settings",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        pageLayoutView: { type: "string", description: "'print', 'normal', 'pageBreakPreview'" },
        zoom: { type: "number", description: "Zoom percentage" },
        gridlines: { type: "boolean", description: "Whether to display gridlines" },
        headers: { type: "boolean", description: "Whether to display row and column headers" },
        showFormulas: { type: "boolean", description: "Whether to display formulas instead of values" },
        showHeadings: { type: "boolean", description: "Whether to display row and column headings" },
        position: { type: "number", description: "Index of the sheet in the whole workbook (0 based)" },
        enableCalculation: { type: "boolean", description: "Whether to enable calculation" },
        visibility: { type: "boolean", description: "Whether to make the sheet visible" }
      },
      required: ["sheet", "pageLayoutView", "zoom", "gridlines", "headers", "showFormulas", "showHeadings", "position", "enableCalculation", "visibility"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "format_chart",
    description: "Format a chart",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Sheet name" },
        chart: { type: "string", description: "Chart name" },
        title: { type: "string", description: "Chart title" },
        type: { type: "string", description: "Chart type" },
        dataSource: { type: "string", description: "Chart data source cell range address" },
        legend: { type: "boolean", description: "Chart legend" },
        axis: { type: "boolean", description: "Chart axis" },
        series: { type: "string", description: "Chart series" },
        dataLabels: { type: "boolean", description: "Chart data labels" },
        width: { type: "number", description: "Chart width" },
        height: { type: "number", description: "Chart height" },
        left: { type: "number", description: "Chart left position" },
        top: { type: "number", description: "Chart top position" },
        fillColor: { type: "string", description: "Chart fill color" },
        borderVisible: { type: "boolean", description: "Chart border visibility" },
        borderColor: { type: "string", description: "Chart border color" },
        borderWidth: { type: "number", description: "Chart border width" },
        borderStyle: { type: "string", description: "Chart border style" },
        borderDashStyle: { type: "string", description: "Chart border dash style" },
        titleVisible: { type: "boolean", description: "Whether title is visible" },
        titleFontName: { type: "string", description: "Title font name" },
        titleFontSize: { type: "number", description: "Title font size" },
        titleFontStyle: { type: "string", description: "Title font style" },
        titleFontBold: { type: "boolean", description: "Title font bold" },
        titleFontItalic: { type: "boolean", description: "Title font italic" },
        titleFontColor: { type: "string", description: "Title font color" },
        titleFormat: { type: "string", description: "Title format" },
        legendVisible: { type: "boolean", description: "Whether legend is visible" },
        legendFontName: { type: "string", description: "Legend font name" },
        legendFontSize: { type: "number", description: "Legend font size" },
        legendFontStyle: { type: "string", description: "Legend font style" },
        legendFontBold: { type: "boolean", description: "Legend font bold" },
        legendFontItalic: { type: "boolean", description: "Legend font italic" },
        legendFontColor: { type: "string", description: "Legend font color" },
        legendFormat: { type: "string", description: "Legend format" },
        axisVisible: { type: "boolean", description: "Whether axis is visible" },
        axisFontName: { type: "string", description: "Axis font name" },
        axisFontSize: { type: "number", description: "Axis font size" },
        axisFontStyle: { type: "string", description: "Axis font style" },
        axisFontBold: { type: "boolean", description: "Axis font bold" },
        axisFontItalic: { type: "boolean", description: "Axis font italic" },
        axisFontColor: { type: "string", description: "Axis font color" },
        axisFormat: { type: "string", description: "Axis format" },
        seriesVisible: { type: "boolean", description: "Whether series is visible" },
        seriesFontName: { type: "string", description: "Series font name" },
        seriesFontSize: { type: "number", description: "Series font size" },
        seriesFontStyle: { type: "string", description: "Series font style" },
        seriesFontBold: { type: "boolean", description: "Series font bold" },
        seriesFontItalic: { type: "boolean", description: "Series font italic" },
        seriesFontColor: { type: "string", description: "Series font color" },
        seriesFormat: { type: "string", description: "Series format" },
        dataLabelsVisible: { type: "boolean", description: "Whether data labels are visible" },
        dataLabelsFontName: { type: "string", description: "Data labels font name" },
        dataLabelsFontSize: { type: "number", description: "Data labels font size" },
        dataLabelsFontStyle: { type: "string", description: "Data labels font style" },
        dataLabelsFontBold: { type: "boolean", description: "Data labels font bold" },
        dataLabelsFontItalic: { type: "boolean", description: "Data labels font italic" },
        dataLabelsFontColor: { type: "string", description: "Data labels font color" },
        dataLabelsFormat: { type: "string", description: "Data labels format" }
      },
      required: ["sheet", "chart", "title", "type", "dataSource", "legend", "axis", "series", "dataLabels", "width", "height", "left", "top", "fillColor", "borderVisible", "borderColor", "borderWidth", "borderStyle", "borderDashStyle", "titleVisible", "titleFontName", "titleFontSize", "titleFontStyle", "titleFontBold", "titleFontItalic", "titleFontColor", "titleFormat", "legendVisible", "legendFontName", "legendFontSize", "legendFontStyle", "legendFontBold", "legendFontItalic", "legendFontColor", "legendFormat", "axisVisible", "axisFontName", "axisFontSize", "axisFontStyle", "axisFontBold", "axisFontItalic", "axisFontColor", "axisFormat", "seriesVisible", "seriesFontName", "seriesFontSize", "seriesFontStyle", "seriesFontBold", "seriesFontItalic", "seriesFontColor", "seriesFormat", "dataLabelsVisible", "dataLabelsFontName", "dataLabelsFontSize", "dataLabelsFontStyle", "dataLabelsFontBold", "dataLabelsFontItalic", "dataLabelsFontColor", "dataLabelsFormat"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "set_calculation_options",
    description: "Set calculation options",
    parameters: {
      type: "object",
      properties: {
        calculationMode: { type: "string", description: "Calculation mode (e.g. 'auto', 'manual')" },
        iterative: { type: "boolean", description: "Whether to enable iterative calculation" },
        maxIterations: { type: "number", description: "Maximum number of iterations for iterative calculation" },
        maxChange: { type: "number", description: "Maximum change for iterative calculation" },
        calculate: { type: "boolean", description: "Whether to calculate the workbook" },
        calculationType: { type: "string", description: "Calculation type (e.g. 'full', 'full_recalculate', 'recalculate')" }
      },
      required: ["calculationMode", "iterative", "maxIterations", "maxChange", "calculate", "calculationType"],
      additionalProperties: false
    },
    strict: true
  },
  {
    type: "function",
    name: "recalculate_ranges",
    description: "Recalculate ranges",
    parameters: {
      type: "object",
      properties: {
        recalculateAll: { type: "boolean", description: "Whether to recalculate all sheets" },
        sheets: { type: "array", items: { type: "string" }, description: "List of sheet names to recalculate" },
        ranges: { type: "array", items: { type: "string" }, description: "List of cell range addresses to recalculate" }
      },
      required: ["recalculateAll", "sheets", "ranges"],
      additionalProperties: false
    },
    strict: true
  }
];