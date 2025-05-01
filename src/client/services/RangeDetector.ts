/**
 * Range Detector for identifying and analyzing key ranges within Excel sheets
 * This is part of the Phase 1 granular workbook state system
 */
import { RangeDetectionResult, RangeInfo, RangeType, createRangeChunk, formatRangeId } from '../models/RangeModels';
import { MetadataChunk, SheetState } from '../models/CommandModels';

/**
 * Service to detect and analyze important ranges within sheets
 */
export class RangeDetector {
  // Minimum dimensions for considering a region significant
  private readonly MIN_REGION_ROWS = 2;
  private readonly MIN_REGION_COLS = 2;
  private readonly MIN_DENSITY_THRESHOLD = 0.4; // 40% density
  /**
   * Detect important ranges within a sheet
   * @param sheet Sheet state to analyze
   * @returns Result with detected ranges and associated chunk IDs
   */
  public detectRanges(sheet: SheetState): RangeDetectionResult {
    const ranges: RangeInfo[] = [];
    const rangeIdToChunkId = new Map<string, string>();
    
    // Skip detection if the sheet has no values
    if (!sheet || !sheet.values || !Array.isArray(sheet.values)) {
      return { ranges, rangeIdToChunkId };
    }
    
    try {
      // Detect tables
      this.detectTables(sheet, ranges);
      
      // Detect named ranges
      this.detectNamedRanges(sheet, ranges);
      
      // Detect formula regions
      this.detectFormulaRegions(sheet, ranges);
      
      // Detect key data regions
      this.detectKeyDataRegions(sheet, ranges);
      
      // Generate chunk IDs for each range
      for (const range of ranges) {
        const chunkId = formatRangeId(range);
        rangeIdToChunkId.set(range.range, chunkId);
      }
      
      return { ranges, rangeIdToChunkId };
    } catch (error) {
      console.error(`Error detecting ranges in sheet ${sheet.name}:`, error);
      return { ranges, rangeIdToChunkId };
    }
  }
  
  /**
   * Generate metadata chunks for all detected ranges
   * @param sheet Sheet state
   * @param detectionResult Range detection result
   * @returns Array of metadata chunks for the ranges
   */
  public createRangeChunks(sheet: SheetState, detectionResult: RangeDetectionResult): MetadataChunk[] {
    const chunks: MetadataChunk[] = [];
    
    for (const range of detectionResult.ranges) {
      try {
        // Extract values and formulas for this range
        const rangeValues = this.extractRangeValues(sheet, range);
        const rangeFormulas = this.extractRangeFormulas(sheet, range);
        
        // Create a chunk for this range
        const chunk = createRangeChunk(range, rangeValues, rangeFormulas);
        chunks.push(chunk);
      } catch (error) {
        console.error(`Error creating chunk for range ${range.range} in sheet ${range.sheetName}:`, error);
      }
    }
    
    return chunks;
  }
  
  /**
   * Detect tables in a sheet (Excel tables, data regions with header rows)
   * @param sheet Sheet state
   * @param ranges Array to populate with detected ranges
   */
  private detectTables(sheet: SheetState, ranges: RangeInfo[]): void {
    // Add any defined Excel tables
    if (sheet.tables && Array.isArray(sheet.tables)) {
      for (const table of sheet.tables) {
        if (!table.range || !table.name) continue;
        
        // Parse the range to get row and column counts
        const { rowCount, columnCount } = this.parseRangeDimensions(table.range);
        
        // Create a range info object for this table
        const rangeInfo: RangeInfo = {
          sheetName: sheet.name,
          range: table.range,
          type: RangeType.DataTable,
          name: table.name,
          headerRow: true, // Assume Excel tables have headers
          totalRow: false, // We don't know this without more analysis
          importance: 90, // Tables are typically important
          rowCount,
          columnCount,
          description: `Excel table with ${table.headers ? table.headers.join(', ') : 'unknown headers'}`
        };
        
        ranges.push(rangeInfo);
      }
    }
    
    // TODO: Add detection of table-like structures that aren't defined as Excel tables
  }
  
  /**
   * Detect named ranges in a sheet
   * @param _sheet Sheet state
   * @param _ranges Array to populate with detected ranges
   */
  private detectNamedRanges(_sheet: SheetState, _ranges: RangeInfo[]): void {
    // If named ranges are available in the sheet state, process them
    if (_sheet.namedRanges && Array.isArray(_sheet.namedRanges)) {
      for (const namedRange of _sheet.namedRanges) {
        if (!namedRange.value || !namedRange.name) continue;
        
        try {
          // Parse A1 notation from the value (typically in format SheetName!A1:B10)
          const match = namedRange.value.match(/!([A-Z]+\d+:[A-Z]+\d+)$/);
          if (!match || !match[1]) continue;
          
          const rangeA1 = match[1];
          const { rowCount, columnCount } = this.parseRangeDimensions(rangeA1);
          
          const rangeInfo: RangeInfo = {
            sheetName: _sheet.name,
            range: rangeA1,
            type: RangeType.NamedRange,
            name: namedRange.name,
            importance: 85, // Named ranges are usually important
            rowCount,
            columnCount,
            description: `Named range: ${namedRange.name}`
          };
          
          _ranges.push(rangeInfo);
        } catch (err) {
          console.warn(`Error parsing named range ${namedRange.name}:`, err);
        }
      }
    }
  }
  
  /**
   * Detect regions containing formulas
   * @param sheet Sheet state
   * @param ranges Array to populate with detected ranges
   */
  private detectFormulaRegions(sheet: SheetState, ranges: RangeInfo[]): void {
    if (!sheet.formulas || !Array.isArray(sheet.formulas)) {
      return;
    }
    
    // Simple implementation: find contiguous blocks of formulas
    const formulaMap: boolean[][] = [];
    
    // Initialize the formula map
    for (let row = 0; row < sheet.formulas.length; row++) {
      formulaMap[row] = [];
      if (!sheet.formulas[row]) continue;
      
      for (let col = 0; col < sheet.formulas[row].length; col++) {
        const hasFormula = !!sheet.formulas[row][col] && 
                          typeof sheet.formulas[row][col] === 'string' && 
                          sheet.formulas[row][col].startsWith('=');
        formulaMap[row][col] = hasFormula;
      }
    }
    
    // Find regions with high formula density
    // Use the same approach as finding data regions, but for formulas
    this.findFormulaRegions(sheet, formulaMap, ranges);
  }
  
  /**
   * Create a range from a detected formula region
   * @param sheet Sheet state
   * @param ranges Array to populate
   * @param start Start coordinates [row, col]
   * @param end End coordinates [row, col]
   * @param formulaCount Number of formulas in the region
   */
  private finalizeFormulaRegion(
    sheet: SheetState, 
    ranges: RangeInfo[], 
    start: [number, number], 
    end: [number, number],
    formulaCount: number
  ): void {
    // Skip tiny regions (1-2 formulas)
    if (formulaCount < 3) {
      return;
    }
    
    // Convert to A1 notation
    const startCol = this.columnToLetter(start[1]);
    const endCol = this.columnToLetter(end[1]);
    const rangeA1 = `${startCol}${start[0] + 1}:${endCol}${end[0] + 1}`;
    
    // Calculate dimensions
    const rowCount = end[0] - start[0] + 1;
    const columnCount = end[1] - start[1] + 1;
    
    // Create the range info
    const rangeInfo: RangeInfo = {
      sheetName: sheet.name,
      range: rangeA1,
      type: RangeType.FormulaRange,
      importance: Math.min(70, 40 + (formulaCount / 2)), // Importance based on formula count
      rowCount,
      columnCount,
      description: `Region with ${formulaCount} formulas`
    };
    
    ranges.push(rangeInfo);
  }
  
  /**
   * Detect key data regions within a sheet
   * @param sheet Sheet state
   * @param ranges Array to populate with detected ranges
   */
  private detectKeyDataRegions(sheet: SheetState, ranges: RangeInfo[]): void {
    if (!sheet.values || !Array.isArray(sheet.values) || sheet.values.length === 0) {
      return;
    }
    
    // If the used range is completely empty, skip detection
    if (sheet.usedRange.rowCount === 0 || sheet.usedRange.columnCount === 0) {
      return;
    }
    
    // Capture the entire used range as a key region
    const usedRangeA1 = this.convertToA1Notation(0, 0, 
                                              sheet.usedRange.rowCount - 1, 
                                              sheet.usedRange.columnCount - 1);
    
    const usedRangeInfo: RangeInfo = {
      sheetName: sheet.name,
      range: usedRangeA1,
      type: RangeType.KeyRegion,
      name: `${sheet.name}_UsedRange`,
      importance: 75,
      rowCount: sheet.usedRange.rowCount,
      columnCount: sheet.usedRange.columnCount,
      description: `Entire used range of sheet ${sheet.name}`
    };
    
    ranges.push(usedRangeInfo);
    
    // Find dense data regions within the sheet using a sliding window approach
    this.findDenseDataRegions(sheet, ranges);
    if (!sheet.formulas || !Array.isArray(sheet.formulas)) {
      return;
    }
    
    // Simple implementation: find contiguous blocks of formulas
  }
  
  /**
   * Find dense data regions within a sheet using a windowing algorithm
   * @param sheet Sheet state to analyze
   * @param ranges Array to populate with detected ranges
   */
  private findDenseDataRegions(sheet: SheetState, ranges: RangeInfo[]): void {
    const values = sheet.values;
    if (!values || values.length === 0) return;
    
    const rowCount = values.length;
    const colCount = Math.max(...values.map(row => row ? row.length : 0));
    
    // Skip very small sheets
    if (rowCount < this.MIN_REGION_ROWS || colCount < this.MIN_REGION_COLS) return;
    
    // Create a density map (cells with non-empty values)
    const densityMap: boolean[][] = [];
    
    for (let r = 0; r < rowCount; r++) {
      densityMap[r] = [];
      for (let c = 0; c < colCount; c++) {
        const hasValue = values[r] && values[r][c] !== null && values[r][c] !== undefined && values[r][c] !== '';
        densityMap[r][c] = hasValue;
      }
    }
    
    // Track regions we've already included
    const processedCells = new Set<string>();
    
    // Look for regions with high density
    for (let startRow = 0; startRow < rowCount - this.MIN_REGION_ROWS; startRow++) {
      for (let startCol = 0; startCol < colCount - this.MIN_REGION_COLS; startCol++) {
        // Skip if this cell is already part of a detected region
        if (processedCells.has(`${startRow},${startCol}`)) continue;
        
        // Find the largest rectangle of data from this starting point
        const region = this.findLargestRegion(densityMap, startRow, startCol, rowCount, colCount);
        
        if (region) {
          const { endRow, endCol, density } = region;
          
          // Only consider regions that meet minimum size and density requirements
          if (endRow - startRow + 1 >= this.MIN_REGION_ROWS && 
              endCol - startCol + 1 >= this.MIN_REGION_COLS && 
              density >= this.MIN_DENSITY_THRESHOLD) {
            
            // Mark all cells in this region as processed
            for (let r = startRow; r <= endRow; r++) {
              for (let c = startCol; c <= endCol; c++) {
                processedCells.add(`${r},${c}`);
              }
            }
            
            // Create range info
            const rangeA1 = this.convertToA1Notation(startRow, startCol, endRow, endCol);
            const rangeInfo: RangeInfo = {
              sheetName: sheet.name,
              range: rangeA1,
              type: RangeType.KeyRegion,
              name: `${sheet.name}_DataRegion_${ranges.length}`,
              importance: Math.min(95, 50 + Math.floor(density * 50)), // Higher density = higher importance
              rowCount: endRow - startRow + 1,
              columnCount: endCol - startCol + 1,
              description: `Data region with ${Math.round(density * 100)}% cell density`
            };
            
            ranges.push(rangeInfo);
          }
        }
      }
    }
  }
  
  /**
   * Find formula regions within a sheet
   * @param sheet Sheet state
   * @param formulaMap Map indicating which cells have formulas
   * @param ranges Array to populate with detected ranges
   */
  private findFormulaRegions(sheet: SheetState, formulaMap: boolean[][], ranges: RangeInfo[]): void {
    const rowCount = formulaMap.length;
    if (rowCount === 0) return;
    
    const colCount = Math.max(...formulaMap.map(row => row ? row.length : 0));
    
    // Skip very small sheets
    if (rowCount < this.MIN_REGION_ROWS || colCount < this.MIN_REGION_COLS) return;
    
    // Track regions we've already included
    const processedCells = new Set<string>();
    
    // Look for regions with formulas
    for (let startRow = 0; startRow < rowCount - this.MIN_REGION_ROWS; startRow++) {
      for (let startCol = 0; startCol < colCount - this.MIN_REGION_COLS; startCol++) {
        // Skip if this cell is already part of a detected region
        if (processedCells.has(`${startRow},${startCol}`)) continue;
        
        // Find the largest rectangle with formulas from this starting point
        const region = this.findLargestRegion(formulaMap, startRow, startCol, rowCount, colCount);
        
        if (region) {
          const { endRow, endCol, density } = region;
          
          // Only consider regions that meet minimum size and density requirements
          if (endRow - startRow + 1 >= this.MIN_REGION_ROWS && 
              endCol - startCol + 1 >= this.MIN_REGION_COLS && 
              density >= this.MIN_DENSITY_THRESHOLD) {
            
            // Mark all cells in this region as processed
            for (let r = startRow; r <= endRow; r++) {
              for (let c = startCol; c <= endCol; c++) {
                processedCells.add(`${r},${c}`);
              }
            }
            
            // Create range info
            const rangeA1 = this.convertToA1Notation(startRow, startCol, endRow, endCol);
            const rangeInfo: RangeInfo = {
              sheetName: sheet.name,
              range: rangeA1,
              type: RangeType.FormulaRange,
              name: `${sheet.name}_FormulaRegion_${ranges.length}`,
              importance: Math.min(90, 50 + Math.floor(density * 50)), // Higher density = higher importance
              rowCount: endRow - startRow + 1,
              columnCount: endCol - startCol + 1,
              description: `Formula region with ${Math.round(density * 100)}% formula density`
            };
            
            ranges.push(rangeInfo);
          }
        }
      }
    }
  }
  
  /**
   * Find the largest contiguous region in a boolean map starting from a given cell
   * @param cellMap Map of which cells have the property (data or formulas)
   * @param startRow Starting row
   * @param startCol Starting column
   * @param maxRows Maximum rows in the sheet
   * @param maxCols Maximum columns in the sheet
   * @returns Region information or null if no significant region found
   */
  private findLargestRegion(
    cellMap: boolean[][], 
    startRow: number, 
    startCol: number, 
    maxRows: number, 
    maxCols: number
  ): { endRow: number, endCol: number, density: number } | null {
    // If the starting cell is empty, no region to find
    if (!cellMap[startRow] || !cellMap[startRow][startCol]) return null;
    
    // Find the maximum possible extents
    let endRow = startRow;
    let endCol = startCol;
    
    // Expand right as far as possible with minimum threshold
    for (let c = startCol + 1; c < maxCols; c++) {
      // Check the column density
      let colDensity = 0;
      for (let r = startRow; r <= endRow; r++) {
        if (cellMap[r] && cellMap[r][c]) colDensity++;
      }
      
      if (colDensity / (endRow - startRow + 1) >= 0.3) {
        endCol = c;
      } else {
        break;
      }
    }
    
    // Expand down as far as possible with minimum threshold
    for (let r = startRow + 1; r < maxRows; r++) {
      // Check the row density
      let rowDensity = 0;
      for (let c = startCol; c <= endCol; c++) {
        if (cellMap[r] && cellMap[r][c]) rowDensity++;
      }
      
      if (rowDensity / (endCol - startCol + 1) >= 0.3) {
        endRow = r;
      } else {
        break;
      }
    }
    
    // Calculate overall density
    let filledCells = 0;
    const totalCells = (endRow - startRow + 1) * (endCol - startCol + 1);
    
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (cellMap[r] && cellMap[r][c]) filledCells++;
      }
    }
    
    const density = filledCells / totalCells;
    
    return { endRow, endCol, density };
  }

  /**
   * Convert zero-based row and column indices to A1 notation
   * @param startRow Start row (0-based)
   * @param startCol Start column (0-based)
   * @param endRow End row (0-based)
   * @param endCol End column (0-based)
   * @returns Range in A1 notation
   */
  private convertToA1Notation(startRow: number, startCol: number, endRow: number, endCol: number): string {
    const startColA1 = this.columnIndexToA1(startCol);
    const endColA1 = this.columnIndexToA1(endCol);
    
    // Excel is 1-based for rows
    return `${startColA1}${startRow + 1}:${endColA1}${endRow + 1}`;
  }
  
  /**
   * Convert 0-based column index to A1 notation
   * @param colIndex 0-based column index
   * @returns Column letter(s) in A1 notation
   */
  private columnIndexToA1(colIndex: number): string {
    let a1 = '';
    let tempIndex = colIndex;
    
    do {
      const remainder = tempIndex % 26;
      a1 = String.fromCharCode(65 + remainder) + a1;
      tempIndex = Math.floor(tempIndex / 26) - 1;
    } while (tempIndex >= 0);
    
    return a1;
  }

  
  /**
   * Convert column letter(s) to 0-based index
   * @param colA1 Column letter(s) in A1 notation (e.g., 'A', 'BC')
   * @returns 0-based column index
   */
  private columnA1ToIndex(colA1: string): number {
    let index = 0;
    for (let i = 0; i < colA1.length; i++) {
      index = index * 26 + colA1.charCodeAt(i) - 64; // 'A' is 65 in ASCII
    }
    return index - 1; // Convert to 0-based
  }
  
  /**
   * Extract values for a specific range
   * @param sheet Sheet state
   * @param range Range information
   * @returns 2D array of values for the range
   */
  private extractRangeValues(sheet: SheetState, range: RangeInfo): any[][] {
    if (!sheet.values) {
      return [];
    }
    
    try {
      // Parse the range coordinates
      const coordinates = this.parseRangeCoordinates(range.range);
      if (!coordinates) return [];
      
      const { startRow, startCol, endRow, endCol } = coordinates;
      
      // Extract the values
      const values: any[][] = [];
      
      for (let row = startRow; row <= endRow; row++) {
        if (!sheet.values[row]) continue;
        
        const rowValues: any[] = [];
        for (let col = startCol; col <= endCol; col++) {
          rowValues.push(sheet.values[row][col]);
        }
        
        values.push(rowValues);
      }
      
      return values;
    } catch (error) {
      console.error(`Error extracting values for range ${range.range}:`, error);
      return [];
    }
  }
  
  /**
   * Extract formulas for a specific range
   * @param sheet Sheet state
   * @param range Range info
   * @returns 2D array of formulas in the range
   */
  private extractRangeFormulas(sheet: SheetState, range: RangeInfo): string[][] {
    if (!sheet.formulas) {
      return [];
    }
    
    try {
      // Parse the range coordinates
      const coordinates = this.parseRangeCoordinates(range.range);
      if (!coordinates) return [];
      
      const { startRow, startCol, endRow, endCol } = coordinates;
      
      // Extract the formulas
      const formulas: string[][] = [];
      
      for (let row = startRow; row <= endRow; row++) {
        if (!sheet.formulas[row]) continue;
        
        const rowFormulas: string[] = [];
        for (let col = startCol; col <= endCol; col++) {
          rowFormulas.push(sheet.formulas[row][col] || '');
        }
        
        formulas.push(rowFormulas);
      }
      
      return formulas;
    } catch (error) {
      console.error(`Error extracting formulas for range ${range.range}:`, error);
      return [];
    }
  }
  
  /**
   * Parse an A1 notation range into row and column indices
   * @param rangeA1 A1 notation range (e.g., "A1:C10")
   * @returns Object with start and end row/column indices (0-based)
   */
  private parseRangeCoordinates(rangeA1: string): { startRow: number, startCol: number, endRow: number, endCol: number } | null {
    try {
      // Split into start and end cell references
      const [startCell, endCell] = rangeA1.split(':');
      
      if (!startCell) return null;
      
      // If no end cell, it's a single cell range
      const endCellRef = endCell || startCell;
      
      // Parse start cell
      const startMatch = startCell.match(/([A-Z]+)([0-9]+)/);
      if (!startMatch) return null;
      
      const startCol = this.letterToColumn(startMatch[1]);
      const startRow = parseInt(startMatch[2], 10) - 1; // Convert to 0-based
      
      // Parse end cell
      const endMatch = endCellRef.match(/([A-Z]+)([0-9]+)/);
      if (!endMatch) return null;
      
      const endCol = this.letterToColumn(endMatch[1]);
      const endRow = parseInt(endMatch[2], 10) - 1; // Convert to 0-based
      
      return { startRow, startCol, endRow, endCol };
    } catch (error) {
      console.error(`Error parsing range coordinates for ${rangeA1}:`, error);
      return null;
    }
  }
  
  /**
   * Parse an A1 notation range to get dimensions
   * @param rangeA1 A1 notation range (e.g., "A1:C10")
   * @returns Object with row and column counts
   */
  private parseRangeDimensions(rangeA1: string): { rowCount: number, columnCount: number } {
    const coords = this.parseRangeCoordinates(rangeA1);
    
    if (!coords) {
      return { rowCount: 0, columnCount: 0 };
    }
    
    return {
      rowCount: coords.endRow - coords.startRow + 1,
      columnCount: coords.endCol - coords.startCol + 1
    };
  }
  
  /**
   * Convert a column index to letter notation (0 -> A, 1 -> B, etc.)
   * @param column Column index (0-based)
   * @returns Column letter(s)
   */
  private columnToLetter(column: number): string {
    let temp: number;
    let letter = '';
    
    while (column >= 0) {
      temp = column % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column / 26) - 1;
    }
    
    return letter;
  }
  
  /**
   * Convert a column letter to index (A -> 0, B -> 1, etc.)
   * @param columnLetter Column letter(s)
   * @returns Column index (0-based)
   */
  private letterToColumn(columnLetter: string): number {
    let column = 0;
    const length = columnLetter.length;
    
    for (let i = 0; i < length; i++) {
      column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    
    return column - 1; // Convert to 0-based
  }
}
