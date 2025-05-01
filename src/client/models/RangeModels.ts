/**
 * Models for range-level granularity in workbook state
 */
import { MetadataChunk } from './CommandModels';

/**
 * Types of range chunks we can capture
 */
export enum RangeType {
  DataTable = 'dataTable',       // Data in tabular format
  PivotTable = 'pivotTable',     // Excel pivot table
  NamedRange = 'namedRange',     // Excel named range
  FormulaRange = 'formulaRange', // Range of cells with formulas
  KeyRegion = 'keyRegion'        // Important region identified by analysis
}

/**
 * Model for a range within a sheet
 */
export interface RangeInfo {
  sheetName: string;          // Name of the sheet containing the range
  range: string;              // A1 notation of the range (e.g., "A1:B10")
  type: RangeType;            // Type of the range
  name?: string;              // Optional name for the range
  headerRow?: boolean;        // Whether the range has a header row
  totalRow?: boolean;         // Whether the range has a total row
  importance: number;         // Estimated importance (0-100)
  rowCount: number;           // Number of rows in the range
  columnCount: number;        // Number of columns in the range
  description?: string;       // Optional description of what the range represents
}

/**
 * Model for range detection results
 */
export interface RangeDetectionResult {
  ranges: RangeInfo[];            // Detected ranges
  rangeIdToChunkId: Map<string, string>; // Mapping from range IDs to chunk IDs
}

/**
 * Format a RangeInfo into a chunk ID
 * @param range The range info object
 * @returns A formatted chunk ID
 */
export function formatRangeId(range: RangeInfo): string {
  const baseId = `Range:${range.sheetName}!${range.range}`;
  
  // Add type suffix for clarity
  return `${baseId}:${range.type}`;
}

/**
 * Create a metadata chunk from a range info
 * @param range The range info object
 * @param values The actual values in the range (optional)
 * @param formulas The actual formulas in the range (optional)
 * @returns A metadata chunk representing the range
 */
export function createRangeChunk(
  range: RangeInfo, 
  values?: any[][], 
  formulas?: string[][]
): MetadataChunk {
  // Generate a unique ID for this range
  const chunkId = formatRangeId(range);
  
  // Create a summary based on the range information
  const summary = generateRangeSummary(range);
  
  // Build the payload
  const payload: any = {
    ...range,
    summary
  };
  
  // Add values and formulas if provided
  if (values) {
    payload.values = values;
  }
  
  if (formulas) {
    payload.formulas = formulas;
  }
  
  // Build dependencies (initially just the parent sheet)
  const refs = [`Sheet:${range.sheetName}`];
  
  return {
    id: chunkId,
    type: 'range',
    etag: new Date().getTime().toString(), // Simple timestamp as etag for now
    payload,
    summary,
    refs,
    lastCaptured: new Date()
  };
}

/**
 * Generate a human-readable summary for a range
 * @param range The range info object
 * @returns A descriptive summary
 */
function generateRangeSummary(range: RangeInfo): string {
  let description = range.name ? `"${range.name}"` : `unnamed ${range.type}`;
  description += ` on sheet "${range.sheetName}" at ${range.range}`;
  
  if (range.headerRow) {
    description += ' with header row';
  }
  
  if (range.totalRow) {
    description += ' with totals row';
  }
  
  description += ` (${range.rowCount}x${range.columnCount})`;
  
  if (range.description) {
    description += `: ${range.description}`;
  }
  
  return description;
}
