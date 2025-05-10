import { WorkbookState, SheetState, CompressedWorkbook, CompressedSheet } from '../../models/CommandModels';

/**
 * Client-side spreadsheet compressor for efficient LLM processing
 */
export class ClientSpreadsheetCompressor {
  /**
   * Compress a workbook state for efficient LLM processing
   * @param workbookState The workbook state to compress
   * @returns Compressed workbook representation
   */
  public compress(workbookState: WorkbookState): CompressedWorkbook {
    // Initialize the compressed workbook
    const compressedWorkbook: CompressedWorkbook = {
      sheets: [],
      activeSheet: workbookState.activeSheet,
      metrics: {
        totalSheets: 0,
        totalCells: 0,
        totalFormulas: 0,
        totalTables: 0,
        totalCharts: 0
      }
    };

    // Process each sheet
    for (const sheet of workbookState.sheets) {
      // Include all sheets, even if they don't have values
      // Create a compressed representation of the sheet
      const compressedSheet = this.compressSheet(sheet);
      compressedWorkbook.sheets.push(compressedSheet);

      // Update workbook metrics
      compressedWorkbook.metrics.totalSheets++;
      compressedWorkbook.metrics.totalCells += compressedSheet.metrics.valueCount + compressedSheet.metrics.formulaCount + compressedSheet.metrics.emptyCount;
      compressedWorkbook.metrics.totalFormulas += compressedSheet.metrics.formulaCount;
      compressedWorkbook.metrics.totalTables += sheet.tables?.length || 0;
      compressedWorkbook.metrics.totalCharts += sheet.charts?.length || 0;
    }

    // Detect the model type
    compressedWorkbook.modelType = this.detectModelType(workbookState);

    // Extract cross-sheet references
    compressedWorkbook.dependencyGraph = this.extractDependencyGraph(workbookState);

    // Detect cell types by color
    compressedWorkbook.colorLegend = this.identifyCellTypesByColor(workbookState);

    return compressedWorkbook;
  }

  /**
   * Compress a sheet state
   * @param sheetState The sheet state to compress
   * @returns Compressed sheet representation
   */
  private compressSheet(sheetState: SheetState): CompressedSheet {
    const compressedSheet: CompressedSheet = {
      name: sheetState.name,
      metrics: {
        rowCount: sheetState.usedRange?.rowCount || 0,
        columnCount: sheetState.usedRange?.columnCount || 0,
        formulaCount: 0,
        valueCount: 0,
        emptyCount: 0
      }
    };

    // Extract key regions
    compressedSheet.keyRegions = this.detectKeyRegions(sheetState);

    // Extract anchors
    compressedSheet.anchors = this.detectAnchors(sheetState);

    // Add tables if available
    if (sheetState.tables && sheetState.tables.length > 0) {
      compressedSheet.tables = sheetState.tables.map(table => ({
        name: table.name,
        range: table.address,
        headers: [] // Would need to extract headers from the actual table
      }));
    }

    // Add charts if available
    if (sheetState.charts && sheetState.charts.length > 0) {
      compressedSheet.charts = sheetState.charts.map(chart => ({
        name: chart.name,
        type: chart.type,
        range: '' // Would need to extract range from the actual chart
      }));
    }

    // Add a summary of the sheet's content
    compressedSheet.summary = this.generateSheetSummary(sheetState);
    
    // Calculate metrics
    if (sheetState.values) {
      for (let i = 0; i < sheetState.values.length; i++) {
        for (let j = 0; j < sheetState.values[i].length; j++) {
          const value = sheetState.values[i][j];
          const formula = sheetState.formulas?.[i]?.[j];

          if (formula && typeof formula === 'string' && formula.startsWith('=')) {
            compressedSheet.metrics.formulaCount++;
          } else if (value !== null && value !== undefined && value !== '') {
            compressedSheet.metrics.valueCount++;
          } else {
            compressedSheet.metrics.emptyCount++;
          }
        }
      }
    } else {
      // Even if there are no values, ensure we provide meaningful information about the sheet
      compressedSheet.metrics.valueCount = 0;
      compressedSheet.metrics.emptyCount = compressedSheet.metrics.rowCount * compressedSheet.metrics.columnCount;
      compressedSheet.metrics.formulaCount = 0;
    }

    return compressedSheet;
  }

  /**
   * Detect key regions in a sheet
   * @param sheetState The sheet state
   * @returns Array of key regions
   */
  private detectKeyRegions(sheetState: SheetState): { name: string; range: string; description?: string }[] {
    const keyRegions: { name: string; range: string; description?: string }[] = [];

    // This is a simplified implementation
    // In a real implementation, we would analyze the sheet to detect regions like:
    // - Headers
    // - Input areas
    // - Output areas
    // - Calculation areas
    // - Data tables

    // For now, we'll just add the entire used range as a key region
    if (sheetState.usedRange) {
      keyRegions.push({
        name: 'UsedRange',
        range: `A1:${this.columnToLetter(sheetState.usedRange.columnCount)}${sheetState.usedRange.rowCount}`,
        description: 'The entire used range of the sheet'
      });
    }

    return keyRegions;
  }

  /**
   * Generate a meaningful summary of a sheet's content and purpose
   * @param sheetState The sheet state to summarize
   * @returns A descriptive summary string
   */
  private generateSheetSummary(sheetState: SheetState): string {
    // Initialize an array to collect summary parts
    const summaryParts: string[] = [];
    
    // Start with basic sheet info
    summaryParts.push(`Sheet '${sheetState.name}'`);
    
    // Add used range info if available
    if (sheetState.usedRange) {
      const { rowCount, columnCount } = sheetState.usedRange;
      if (rowCount > 0 && columnCount > 0) {
        summaryParts.push(`contains ${rowCount} rows and ${columnCount} columns of data`);
      } else {
        summaryParts.push('appears to be empty');
      }
    } else {
      summaryParts.push('has no data');
    }
    
    // Add information about formulas if available
    if (sheetState.formulas) {
      // Count formulas
      let formulaCount = 0;
      for (const row of sheetState.formulas) {
        for (const cell of row) {
          if (cell && typeof cell === 'string' && cell.startsWith('=')) {
            formulaCount++;
          }
        }
      }
      
      if (formulaCount > 0) {
        summaryParts.push(`contains ${formulaCount} formulas`);
      }
    }
    
    // Add tables info if available
    if (sheetState.tables && sheetState.tables.length > 0) {
      summaryParts.push(`has ${sheetState.tables.length} table${sheetState.tables.length > 1 ? 's' : ''}`);
    }
    
    // Add charts info if available
    if (sheetState.charts && sheetState.charts.length > 0) {
      const chartTypes = [...new Set(sheetState.charts.map(chart => chart.type))];
      summaryParts.push(`contains ${sheetState.charts.length} chart${sheetState.charts.length > 1 ? 's' : ''} (${chartTypes.join(', ')})`);
    }
    
    // Join all parts into a coherent summary
    return summaryParts.join(' ');
  }
  
  /**
   * Detect anchors in a sheet
   * @param sheetState The sheet state
   * @returns Array of anchors
   */
  private detectAnchors(sheetState: SheetState): { cell: string; value: any; type: string }[] {
    const anchors: { cell: string; value: any; type: string }[] = [];

    // This is a simplified implementation
    // In a real implementation, we would analyze the sheet to detect anchors like:
    // - Headers
    // - Labels
    // - Key inputs
    // - Key outputs

    // For now, we'll just look for cells with specific keywords
    if (sheetState.values) {
      for (let i = 0; i < sheetState.values.length; i++) {
        for (let j = 0; j < sheetState.values[i].length; j++) {
          const value = sheetState.values[i][j];
          
          if (typeof value === 'string') {
            const lowerValue = value.toLowerCase();
            
            // Check for common financial terms
            if (lowerValue.includes('revenue') || 
                lowerValue.includes('income') || 
                lowerValue.includes('expense') || 
                lowerValue.includes('profit') || 
                lowerValue.includes('loss') || 
                lowerValue.includes('total') || 
                lowerValue.includes('subtotal')) {
              
              anchors.push({
                cell: `${this.columnToLetter(j + 1)}${i + 1}`,
                value,
                type: 'label'
              });
            }
          }
        }
      }
    }

    return anchors;
  }

  /**
   * Detect the type of financial model
   * @param workbookState The workbook state
   * @returns The detected model type
   */
  private detectModelType(workbookState: WorkbookState): string {
    // This is a simplified implementation
    // In a real implementation, we would analyze the workbook to detect the model type
    
    // Check for common sheet names
    const sheetNames = workbookState.sheets.map(sheet => sheet.name.toLowerCase());
    
    if (sheetNames.some(name => name.includes('income') || name.includes('p&l') || name.includes('profit'))) {
      if (sheetNames.some(name => name.includes('balance'))) {
        if (sheetNames.some(name => name.includes('cash flow'))) {
          return 'Three Statement Financial Model';
        }
        return 'Income Statement and Balance Sheet Model';
      }
      return 'Income Statement Model';
    }
    
    if (sheetNames.some(name => name.includes('dcf') || name.includes('discounted cash flow'))) {
      return 'Discounted Cash Flow Model';
    }
    
    if (sheetNames.some(name => name.includes('lbo') || name.includes('leveraged buyout'))) {
      return 'Leveraged Buyout Model';
    }
    
    if (sheetNames.some(name => name.includes('merger') || name.includes('acquisition') || name.includes('m&a'))) {
      return 'Merger and Acquisition Model';
    }
    
    return 'Generic Financial Model';
  }

  /**
   * Extract the dependency graph from a workbook
   * @param workbookState The workbook state
   * @returns The dependency graph
   */
  private extractDependencyGraph(workbookState: WorkbookState): any {
    // This is a simplified implementation
    // In a real implementation, we would analyze formulas to build a dependency graph
    
    const nodes: string[] = [];
    const edges: { source: string; target: string; type: string }[] = [];
    
    // Add all sheets as nodes
    for (const sheet of workbookState.sheets) {
      nodes.push(sheet.name);
    }
    
    // Look for cross-sheet references in formulas
    for (const sheet of workbookState.sheets) {
      if (!sheet.formulas) continue;
      
      for (let i = 0; i < sheet.formulas.length; i++) {
        for (let j = 0; j < sheet.formulas[i].length; j++) {
          const formula = sheet.formulas[i][j];
          
          if (typeof formula === 'string' && formula.startsWith('=')) {
            // Check for references to other sheets
            for (const otherSheet of workbookState.sheets) {
              if (otherSheet.name !== sheet.name && formula.includes(`${otherSheet.name}!`)) {
                edges.push({
                  source: sheet.name,
                  target: otherSheet.name,
                  type: 'reference'
                });
                break;
              }
            }
          }
        }
      }
    }
    
    return { nodes, edges };
  }

  /**
   * Identify cell types by color
   * @param workbookState The workbook state
   * @returns Array of color legend entries
   */
  private identifyCellTypesByColor(_workbookState: WorkbookState): any[] {
    // This is a simplified implementation
    // In a real implementation, we would analyze cell formats to identify patterns
    
    // For now, we'll just return a default color legend
    return [
      { color: '#D9D9D9', type: 'header' },
      { color: '#E2EFDA', type: 'input' },
      { color: '#FFF2CC', type: 'calculation' },
      { color: '#DDEBF7', type: 'output' }
    ];
  }

  /**
   * Convert a column number to a letter (e.g., 1 -> A, 27 -> AA)
   * @param column The column number (1-based)
   * @returns The column letter
   */
  private columnToLetter(column: number): string {
    let temp: number;
    let letter = '';
    
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    
    return letter;
  }
}
