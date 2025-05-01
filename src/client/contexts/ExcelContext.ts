import * as React from 'react';
import { ReactNode } from 'react';

// Office.js type declarations
declare const Office: any;
declare const Excel: any;

// Define the Excel context interface
export interface ExcelContextType {
  activeCell: {
    sheet: string;
    address: string;
    value?: any;
  } | null;
  selectedRange: {
    sheet: string;
    address: string;
    values?: any[][];
  } | null;
  isExcelConnected: boolean;
  getWorksheetContext: () => Promise<any>;
  populateCell: (address: string, value: any) => Promise<void>;
}

// Create the Excel context with default values
const defaultContext: ExcelContextType = {
  activeCell: null,
  selectedRange: null,
  isExcelConnected: false,
  getWorksheetContext: async () => ({}),
  populateCell: async () => {}
};

// Create the context
const ExcelContext = React.createContext<ExcelContextType>(defaultContext);

// Custom hook to use the Excel context
export const useExcelContext = (): ExcelContextType => {
  return React.useContext(ExcelContext);
};

// Provider component props type
interface ExcelContextProviderProps {
  children: ReactNode;
}

/**
 * Excel Context Provider Component
 * Manages Excel-related state and provides methods for interacting with Excel
 */
export const ExcelContextProvider = (props: ExcelContextProviderProps): JSX.Element => {
  // State for tracking Excel context
  const [activeCell, setActiveCell] = React.useState<ExcelContextType['activeCell']>(null);
  const [selectedRange, setSelectedRange] = React.useState<ExcelContextType['selectedRange']>(null);
  const [isExcelConnected, setIsExcelConnected] = React.useState<boolean>(false);

  /**
   * Get the current worksheet context using Office.js API
   * Returns worksheet name, headers, data, and selected range information
   */
  const getWorksheetContext = async (): Promise<any> => {
    try {
      // Check if Office and Excel are available
      if (typeof Excel === 'undefined') {
        console.warn('Excel API not available');
        return null;
      }

      return await Excel.run(async (context) => {
        // Get active worksheet
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load('name');
        
        // Get used range for headers and data
        const usedRange = worksheet.getUsedRange();
        usedRange.load(['address', 'values', 'rowCount', 'columnCount']);
        
        // Get selected range
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load(['address', 'values', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
        
        await context.sync();
        
        // Extract headers (assuming first row contains headers)
        const headers: string[] = [];
        if (usedRange.rowCount > 0) {
          for (const h of usedRange.values[0]) {
            headers.push(h?.toString() || '');
          }
        }
        
        // Extract data (excluding headers)
        let data: any[][] = [];
        if (usedRange.rowCount > 1) {
          data = usedRange.values.slice(1);
        }
        
        return {
          worksheetName: worksheet.name,
          selectedRange: {
            startRow: selectedRange.rowIndex,
            startColumn: selectedRange.columnIndex,
            rowCount: selectedRange.rowCount,
            columnCount: selectedRange.columnCount,
            address: selectedRange.address
          },
          headers,
          data
        };
      });
    } catch (error) {
      console.error('Error getting worksheet context:', error);
      return null;
    }
  };

  /**
   * Populate a cell with a value using Office.js API
   * @param address - Cell address (e.g., "A1")
   * @param value - Value to set in the cell
   */
  const populateCell = async (address: string, value: any): Promise<void> => {
    try {
      // Check if Office and Excel are available
      if (typeof Excel === 'undefined') {
        console.warn('Excel API not available');
        return;
      }

      await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
        range.values = [[value]];
        await context.sync();
      });
    } catch (error) {
      console.error('Error populating cell:', error);
    }
  };

  // Track selection changes in Excel
  React.useEffect(() => {
    const handleSelectionChanged = async () => {
      try {
        // Check if Office and Excel are available
        if (typeof Excel === 'undefined') {
          return;
        }

        await Excel.run(async (context) => {
          // Get active cell/range
          const range = context.workbook.getSelectedRange();
          range.load(['address', 'values', 'rowCount', 'columnCount']);
          const worksheet = range.worksheet;
          worksheet.load('name');
          
          await context.sync();
          
          // Update active cell if it's a single cell
          if (range.rowCount === 1 && range.columnCount === 1) {
            setActiveCell({
              sheet: worksheet.name,
              address: range.address,
              value: range.values[0][0]
            });
          } else {
            setActiveCell(null);
          }
          
          // Update selected range
          setSelectedRange({
            sheet: worksheet.name,
            address: range.address,
            values: range.values
          });
        });
      } catch (error) {
        console.error('Error handling selection change:', error);
      }
    };

    // Check if Office is available and initialized
    const setupSelectionHandler = () => {
      if (typeof Office !== 'undefined' && Office.context?.document) {
        // Register event handler
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          handleSelectionChanged
        );
        
        // Initial selection check
        handleSelectionChanged();
        return true;
      }
      return false;
    };

    // Try to set up handler, or retry after Office is initialized
    if (!setupSelectionHandler() && typeof Office !== 'undefined') {
      Office.onReady(() => {
        setupSelectionHandler();
      });
    }
    
    // Clean up event handler
    return () => {
      if (typeof Office !== 'undefined' && Office.context?.document) {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          handleSelectionChanged
        );
      }
    };
  }, []);

  // Check if Excel is connected
  React.useEffect(() => {
    const checkExcelConnection = () => {
      // Office.js is available if the add-in is running in Excel
      const isConnected = typeof Office !== 'undefined' && typeof Excel !== 'undefined';
      setIsExcelConnected(isConnected);
    };
    
    checkExcelConnection();
    
    // Set up Office.onReady to detect when Office.js is fully initialized
    if (typeof Office !== 'undefined') {
      Office.onReady(() => {
        checkExcelConnection();
      });
    }
    
    // Periodic check for connection status
    const interval = setInterval(checkExcelConnection, 5000);
    
    return () => clearInterval(interval);
  }, []);

  // Create the context value object
  const contextValue: ExcelContextType = {
    activeCell,
    selectedRange,
    isExcelConnected,
    getWorksheetContext,
    populateCell
  };

  // Return the provider component using React.createElement instead of JSX
  return React.createElement(
    ExcelContext.Provider,
    { value: contextValue },
    props.children
  );
};

export default ExcelContext;
