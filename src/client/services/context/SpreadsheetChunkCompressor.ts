import { 
  CompressedSheet, 
  CompressedWorkbook, 
  MetadataChunk, 
  SheetState, 
  WorkbookMetrics 
} from '../../models/CommandModels';
import * as CryptoJS from 'crypto-js';
import { WorkbookMetadataLogger } from './WorkbookMetadataLogger';


/**
 * Interface for processed sheet data returned by processSheetData
 */
interface ProcessedSheetData {
  metrics: {
    rowCount: number;
    columnCount: number;
    formulaCount: number;
    valueCount: number;
    emptyCount: number;
  };
  anchors?: any[];
  cells?: {
    address: string;   // Cell address in A1 notation
    value: any;        // Cell value
    formula?: string;  // Formula if present
    type: string;      // Data type
    format?: any;      // Format information if available
  }[];
}

/**
 * Compressor for individual workbook chunks (sheets, ranges)
 * This extracts the chunk-based compression logic from ClientSpreadsheetCompressor
 */
export class SpreadsheetChunkCompressor {
  
  /**
   * Compress a sheet to a metadata chunk
   * @param sheet The sheet to compress
   * @returns The metadata chunk
   */
  public compressSheetToChunk(sheet: SheetState): MetadataChunk {
    try {
      if (!sheet) {
        throw new Error('Sheet is undefined');
      }

      const sheetName = sheet.name;
      const compressedSheet = this.compressSheet(sheet);
      
      // Attach full formulas array for dependency analysis (not used for prompt context)
      compressedSheet.formulas = sheet.formulas;
      
      // Calculate ETag (hash) based on sheet content
      const etag = this.calculateETag(sheet);
      
      return {
        id: `Sheet:${sheetName}`,
        type: 'sheet',
        etag,
        payload: compressedSheet,
        summary: compressedSheet.summary,
        refs: [], // Will be populated by dependency analyzer
        lastCaptured: new Date()
      };
    } catch (error) {
      console.error(`Error in compressSheetToChunk for sheet ${sheet.name}:`, error);
      throw new Error(`Failed to compress sheet "${sheet.name}": ${error.message}`);
    }
  }

  /**
   * Compresses a sheet state into a compressed sheet format
   * This is adapted from the original ClientSpreadsheetCompressor.compressSheet method
   * @param sheet The sheet state to compress
   * @returns The compressed sheet
   */
  public compressSheet(sheet: SheetState): CompressedSheet {
    if (!sheet) {
      throw new Error('Sheet is undefined');
    }

    // Basic sheet info
    const compressedSheet: CompressedSheet = {
      name: sheet.name,
      summary: `Sheet ${sheet.name} with ${sheet.values ? sheet.values.length : 0} rows`,
    };
    
    // Process sheet data to extract key information and metrics
    const processedData = this.processSheetData(sheet);
    
    // Add metrics to the compressed sheet
    compressedSheet.metrics = processedData.metrics;
    
    // Add anchors if present
    if (processedData.anchors && processedData.anchors.length > 0) {
      compressedSheet.anchors = processedData.anchors;
    }
    
    // Add detailed cell data if present
    if (processedData.cells && processedData.cells.length > 0) {
      compressedSheet.cells = processedData.cells;
      console.log(`Added ${processedData.cells.length} detailed cell records to sheet ${sheet.name}`);
    }
    
    // Add tables if present
    if (sheet.tables && sheet.tables.length > 0) {
      compressedSheet.tables = sheet.tables.map(table => ({
        name: table.name,
        range: table.range,
        headers: table.headers || []
      }));
    }

    // Add charts if available
    if (sheet.charts && sheet.charts.length > 0) {
      compressedSheet.charts = sheet.charts.map(chart => ({
        name: chart.name,
        type: chart.type,
        range: chart.range
      }));
    }

    // Generate a summary for the sheet
    compressedSheet.summary = this.generateSheetSummary(compressedSheet);

    return compressedSheet;
  }
  
  /**
   * Calculate aggregated workbook metrics from chunks
   * @param chunks Array of metadata chunks to calculate metrics from
   * @returns Aggregated workbook metrics
   */
  public calculateWorkbookMetrics(chunks: MetadataChunk[]): WorkbookMetrics {
    // Initialize metrics
    const metrics: WorkbookMetrics = {
      totalSheets: 0,
      totalCells: 0,
      totalFormulas: 0,
      totalTables: 0,
      totalCharts: 0
    };
    
    // Only count sheet chunks for totalSheets
    const sheetChunks = chunks.filter(chunk => chunk.type === 'sheet');
    metrics.totalSheets = sheetChunks.length;
    
    // Aggregate metrics from all chunks
    for (const chunk of chunks) {
      if (!chunk.payload) continue;
      
      // Get metrics from sheet chunk
      if (chunk.type === 'sheet' && chunk.payload.metrics) {
        const sheetMetrics = chunk.payload.metrics;
        
        // Add cell counts
        if (sheetMetrics.valueCount !== undefined) {
          metrics.totalCells += sheetMetrics.valueCount || 0;
        }
        
        // Add formula counts
        if (sheetMetrics.formulaCount !== undefined) {
          metrics.totalFormulas += sheetMetrics.formulaCount || 0;
        }
        
        // Add tables
        if (chunk.payload.tables && Array.isArray(chunk.payload.tables)) {
          metrics.totalTables += chunk.payload.tables.length;
        }
        
        // Add charts
        if (chunk.payload.charts && Array.isArray(chunk.payload.charts)) {
          metrics.totalCharts += chunk.payload.charts.length;
        }
      }
    }
    
    return metrics;
  }

  /**
   * Generate a summary for a sheet
   * @param sheet The compressed sheet
   * @returns A summary string
   */
  private generateSheetSummary(sheet: CompressedSheet): string {
    let summary = `Sheet ${sheet.name}`;
    
    // Add basic metrics
    if (sheet.metrics) {
      summary += ` with ${sheet.metrics.valueCount || 0} values`;
      
      if (sheet.metrics.formulaCount && sheet.metrics.formulaCount > 0) {
        summary += `, ${sheet.metrics.formulaCount} formulas`;
      }
    }
    
    // Add table information
    if (sheet.tables && sheet.tables.length > 0) {
      summary += `, ${sheet.tables.length} table${sheet.tables.length === 1 ? '' : 's'}`;
    }
    
    // Add chart information
    if (sheet.charts && sheet.charts.length > 0) {
      summary += `, ${sheet.charts.length} chart${sheet.charts.length === 1 ? '' : 's'}`;
    }
    
    return summary;
  }
  
  /**
   * Calculate a hash-based ETag for a sheet
   * @param sheet The sheet to hash
   * @returns A hash string for the sheet content
   */
  private calculateETag(sheet: SheetState): string {
    try {
      // Create a string representation of important sheet content
      const contentParts = [];
      
      // Add sheet name
      contentParts.push(sheet.name);
      
      // Add a sample of values (first 10 rows)
      if (sheet.values && sheet.values.length > 0) {
        for (let i = 0; i < Math.min(10, sheet.values.length); i++) {
          contentParts.push(JSON.stringify(sheet.values[i]));
        }
      }
      
      // Add a sample of formulas (first 10 rows)
      if (sheet.formulas && sheet.formulas.length > 0) {
        for (let i = 0; i < Math.min(10, sheet.formulas.length); i++) {
          contentParts.push(JSON.stringify(sheet.formulas[i]));
        }
      }
      
      // Add tables and charts info
      if (sheet.tables) {
        contentParts.push(JSON.stringify(sheet.tables));
      }
      
      if (sheet.charts) {
        contentParts.push(JSON.stringify(sheet.charts));
      }
      
      // Calculate hash
      const content = contentParts.join('\n');
      return CryptoJS.SHA256(content).toString();
    } catch (error) {
      console.error(`Error calculating ETag for sheet ${sheet.name}:`, error);
      // Use timestamp as fallback ETag
      return new Date().getTime().toString();
    }
  }

  /**
   * Process sheet data to extract key information and metrics
   * @param sheet The sheet data
   * @returns Object with processed metrics and key information
   */
  private processSheetData(sheet: SheetState): ProcessedSheetData {
    const metrics = {
      rowCount: sheet.values ? sheet.values.length : 0,
      columnCount: 0,
      formulaCount: 0,
      valueCount: 0,
      emptyCount: 0
    };
    
    const anchors: any[] = [];
    const cells: {
      address: string;
      value: any;
      formula?: string;
      type: string;
      format?: any;
    }[] = [];
    
    // Process all cells to extract key information
    if (sheet.values && Array.isArray(sheet.values)) {
      for (let row = 0; row < sheet.values.length; row++) {
        const valueRow = sheet.values[row] || [];
        const formulaRow = (sheet.formulas && Array.isArray(sheet.formulas) && sheet.formulas[row]) 
                          ? sheet.formulas[row] 
                          : [];
        const formatRow = (sheet.formats && Array.isArray(sheet.formats) && sheet.formats[row])
                          ? sheet.formats[row]
                          : [];
        
        metrics.columnCount = Math.max(metrics.columnCount, valueRow.length);
        
        for (let col = 0; col < Math.max(valueRow.length, formulaRow.length); col++) {
          try {
            const value = valueRow[col];
            const formula = formulaRow[col];
            const format = formatRow[col];
            const cellAddress = this.columnToLetter(col) + (row + 1);
            
            // Only process formula if it's a valid string
            if (formula && typeof formula === 'string' && formula.trim().startsWith('=')) {
              metrics.formulaCount++;
              
              // Store all formulas in the cells array
              cells.push({
                address: cellAddress,
                value: value,
                formula: formula,
                type: 'formula',
                format: format
              });
              
              // Add important formulas as anchors for backward compatibility
              if (this.isKeyFormula(formula)) {
                anchors.push({
                  cell: cellAddress,
                  value: formula,
                  type: 'formula'
                });
              }
            } else if (value !== undefined && value !== null && value !== '') {
              metrics.valueCount++;
              
              // Store all non-empty cells in the cells array
              const valueType = typeof value;
              const processedValue = valueType === 'object' ? JSON.stringify(value) : value;
              
              cells.push({
                address: cellAddress,
                value: processedValue,
                type: valueType,
                format: format
              });
              
              // Add important values as anchors for backward compatibility
              if (this.isKeyValue(value)) {
                anchors.push({
                  cell: cellAddress,
                  value: processedValue,
                  type: valueType
                });
              }
            } else {
              metrics.emptyCount++;
            }
          } catch (cellError) {
            console.warn(`Error processing cell at row ${row}, col ${col}:`, cellError);
            // Continue with the next cell
          }
        }
      }
    }
    
    // Return both cells and anchors
    const result: ProcessedSheetData = { metrics };
    
    if (cells.length > 0) {
      result.cells = cells;
    }
    
    if (anchors.length > 0) {
      result.anchors = anchors;
    }
    
    return result;
  }
  
  /**
   * Determine if a formula is a key formula (important for context)
   * @param formula The formula to check
   * @returns True if it's an important formula
   */
  private isKeyFormula(formula: string): boolean {
    // Check if formula is a string and not empty
    if (!formula || typeof formula !== 'string') return false;
    
    const importantFormulaPrefixes = [
      '=SUM', '=AVERAGE', '=COUNT', '=MAX', '=MIN', 
      '=IF', '=VLOOKUP', '=HLOOKUP', '=INDEX', '=MATCH',
      '=SUMIF', '=COUNTIF', '=AVERAGEIF',
      '=NPV', '=IRR', '=PMT', '=FV', '=PV',
      '=DATE', '=TODAY', '=NOW',
      '=CONCATENATE', '=FIND', '=SEARCH',
      '=OFFSET', '=INDIRECT'
    ];
    
    const formulaUpper = formula.toUpperCase();
    return importantFormulaPrefixes.some(prefix => formulaUpper.startsWith(prefix));
  }
  
  /**
   * Public wrapper for testing isKeyFormula
   */
  public test_isKeyFormula(formula: string): boolean {
    return this.isKeyFormula(formula);
  }

  /**
   * Determine if a value is a key value (important for context)
   * @param value The value to check
   * @returns True if it's an important value
   */
  private isKeyValue(value: any): boolean {
    // Skip non-string values or empty strings
    if (value === null || value === undefined) return false;
    
    // If it's a string, check for key terms
    if (typeof value === 'string') {
      // Normalize the string for comparison
      const normalizedValue = value.trim().toLowerCase();
      
      // Skip if empty after normalizing
      if (normalizedValue === '') return false;
      
      // Skip short strings unless they look like a key header
      if (normalizedValue.length < 4 && !/^[a-z]+[0-9]*$/i.test(normalizedValue)) return false;
      
      // Check for patterns that suggest this is a header or label
      const keyTerms = ['total', 'sum', 'average', 'count', 'revenue', 'expense', 
                       'income', 'profit', 'loss', 'asset', 'liability', 'equity',
                       'sales', 'cost', 'margin', 'rate', 'growth', 'price',
                       'date', 'year', 'month', 'quarter', 'forecast', 'budget',
                       'actual', 'variance', 'assumption', 'scenario'];
                       
      return keyTerms.some(term => normalizedValue.includes(term));
    }
    
    // For non-string values, include numbers that appear significant
    if (typeof value === 'number') {
      // Include round numbers that might be significant
      if (value % 100 === 0 && value !== 0) return true;
      if (value >= 10000) return true;
    }
    
    return false;
  }

  /**
   * Public wrapper for testing isKeyValue
   */
  public test_isKeyValue(value: any): boolean {
    return this.isKeyValue(value);
  }
  
  /**
   * Convert column index to Excel letter (e.g., 0 -> A, 25 -> Z, 26 -> AA)
   * @param column The column index (0-based)
   * @returns The Excel column letter
   */
  private columnToLetter(column: number): string {
    let temp = column;
    let letter = '';
    
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    
    return letter;
  }
}
