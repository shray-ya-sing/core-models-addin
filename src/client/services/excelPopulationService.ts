/**
 * Service for populating Excel with data from indexed documents
 */
export class ExcelPopulationService {
  /**
   * Populate Excel with data from a suggestion
   * @param suggestion The suggestion to populate
   * @param targetCell The target cell to populate
   * @returns Promise that resolves when the operation is complete
   */
  async populateSuggestion(suggestion: any, targetCell?: { sheet: string; address: string }) {
    try {
      // Get the active cell if no target cell is provided
      let cell = targetCell;
      
      if (!cell) {
        cell = await this.getActiveCell();
      }
      
      if (!cell) {
        throw new Error('No target cell available');
      }
      
      // Populate Excel based on the suggestion type
      if (suggestion.dataPoints && suggestion.dataPoints.length > 0) {
        await this.populateDataPoints(suggestion.dataPoints, cell);
      } else if (suggestion.suggestedRange) {
        await this.populateRange(suggestion, cell);
      } else {
        throw new Error('Unsupported suggestion type');
      }
      
      return {
        success: true,
        message: 'Data populated successfully'
      };
    } catch (error) {
      console.error('Error populating Excel:', error);
      throw error;
    }
  }
  
  /**
   * Get the active cell in Excel
   * @returns Promise that resolves with the active cell
   */
  private async getActiveCell() {
    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(['address']);
        
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load(['name']);
        
        await context.sync();
        
        return {
          sheet: worksheet.name,
          address: range.address.split('!')[1] || range.address
        };
      });
    } catch (error) {
      console.error('Error getting active cell:', error);
      return null;
    }
  }
  
  /**
   * Populate Excel with data points
   * @param dataPoints The data points to populate
   * @param startCell The starting cell
   * @returns Promise that resolves when the operation is complete
   */
  private async populateDataPoints(dataPoints: any[], startCell: { sheet: string; address: string }) {
    try {
      await Excel.run(async (context) => {
        console.log('Populating data points:', JSON.stringify(dataPoints, null, 2));
        console.log('Start cell:', startCell);
        
        const worksheet = context.workbook.worksheets.getItem(startCell.sheet);
        
        // Get the starting cell coordinates
        const range = worksheet.getRange(startCell.address);
        range.load(['rowIndex', 'columnIndex']);
        
        await context.sync();
        
        const startRow = range.rowIndex;
        const startColumn = range.columnIndex;
        console.log(`Starting at row ${startRow}, column ${startColumn}`);
        
        let currentRow = startRow;
        
        // Populate data points
        for (let i = 0; i < dataPoints.length; i++) {
          const dataPoint = dataPoints[i];
          console.log(`Processing data point ${i}:`, dataPoint);
          
          if (dataPoint.type === 'table') {
            console.log('Handling table-type data point');
            
            // Add title as header
            const titleCell = worksheet.getCell(currentRow, startColumn);
            titleCell.values = [[dataPoint.name || 'Table']];
            titleCell.format.font.bold = true;
            currentRow++;
            
            // Get headers and rows from the data point
            const headers = dataPoint.metadata?.headers || [];
            const rows = Array.isArray(dataPoint.value) ? dataPoint.value : [];
            
            console.log('Table headers:', headers);
            console.log('Table rows:', rows);
            
            // Add headers
            if (headers.length > 0) {
              for (let j = 0; j < headers.length; j++) {
                const headerCell = worksheet.getCell(currentRow, startColumn + j);
                headerCell.values = [[headers[j]]];
                headerCell.format.font.bold = true;
              }
              currentRow++;
            }
            
            // Add rows
            for (let j = 0; j < rows.length; j++) {
              const row = rows[j];
              if (Array.isArray(row)) {
                for (let k = 0; k < row.length; k++) {
                  const cell = worksheet.getCell(currentRow, startColumn + k);
                  cell.values = [[row[k]]];
                }
              }
              currentRow++;
            }
            
            // Add source attribution
            if (dataPoint.source) {
              const sourceCell = worksheet.getCell(currentRow, startColumn);
              sourceCell.values = [[`Source: ${dataPoint.source.documentTitle || 'Unknown'}`]];
              sourceCell.format.font.italic = true;
              sourceCell.format.font.size = 9;
              currentRow++;
            }
            
            // Add extra space after the table
            currentRow++;
          } else {
            // Handle regular data points (non-table)
            // Add label in first column
            const labelCell = worksheet.getCell(currentRow, startColumn);
            labelCell.values = [[dataPoint.name || dataPoint.label || `Data ${i + 1}`]];
            
            // Add value in second column
            const valueCell = worksheet.getCell(currentRow, startColumn + 1);
            valueCell.values = [[dataPoint.value]];
            
            currentRow++;
          }
        }
        
        await context.sync();
      });
    } catch (error) {
      console.error('Error populating data points:', error);
      throw error;
    }
  }
  
  /**
   * Populate Excel with a range of data
   * @param suggestion The suggestion containing the range data
   * @param startCell The starting cell
   * @returns Promise that resolves when the operation is complete
   */
  private async populateRange(suggestion: any, startCell: { sheet: string; address: string }) {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(startCell.sheet);
        
        // Get the starting cell coordinates
        const range = worksheet.getRange(startCell.address);
        range.load(['rowIndex', 'columnIndex']);
        
        await context.sync();
        
        const startRow = range.rowIndex;
        const startColumn = range.columnIndex;
        
        // Create a range with the suggested dimensions
        const targetRange = worksheet.getRangeByIndexes(
          startRow,
          startColumn,
          suggestion.suggestedRange.rowCount,
          suggestion.suggestedRange.columnCount
        );
        
        // Populate the range with the data
        if (suggestion.data) {
          targetRange.values = suggestion.data;
        } else if (suggestion.dataPoints && suggestion.dataPoints.length > 0) {
          // Convert data points to a 2D array
          const data = [];
          for (let i = 0; i < suggestion.suggestedRange.rowCount; i++) {
            const row = [];
            for (let j = 0; j < suggestion.suggestedRange.columnCount; j++) {
              const index = i * suggestion.suggestedRange.columnCount + j;
              row.push(index < suggestion.dataPoints.length ? suggestion.dataPoints[index].value : '');
            }
            data.push(row);
          }
          targetRange.values = data;
        }
        
        await context.sync();
      });
    } catch (error) {
      console.error('Error populating range:', error);
      throw error;
    }
  }
}

export default ExcelPopulationService;
