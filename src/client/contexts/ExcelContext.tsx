import * as React from 'react';
import { ReactNode } from 'react';

// Add type declarations for Office.js globals
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

// Create the context with default values
const ExcelContext = React.createContext<ExcelContextType>({
  activeCell: null,
  selectedRange: null,
  isExcelConnected: false,
  getWorksheetContext: async () => ({}),
  populateCell: async () => {}
});

// Custom hook to use the Excel context
export const useExcelContext = (): ExcelContextType => React.useContext(ExcelContext);

// Provider component
export const ExcelContextProvider: React.FC<{children: ReactNode}> = ({ children }) => {
  const [activeCell, setActiveCell] = React.useState<ExcelContextType['activeCell']>(null);
  const [selectedRange, setSelectedRange] = React.useState<ExcelContextType['selectedRange']>(null);
  const [isExcelConnected, setIsExcelConnected] = React.useState<boolean>(false);

  // Get the current worksheet context using Office.js API
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
        let headers: string[] = [];
        if (usedRange.rowCount > 0) {
          headers = usedRange.values[0].map((h: any) => h?.toString() || '');
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

  // Populate a cell with a value using Office.js API
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
    const checkExcelConnection = async () => {
      // First check if Office and Excel objects exist
      if (typeof Office === 'undefined' || typeof Excel === 'undefined') {
        console.log('Office or Excel API not available');
        setIsExcelConnected(false);
        return;
      }
      
      // Then verify we can actually use the Excel API
      try {
        // Try a simple Excel operation to verify the connection
        await Excel.run(async (context) => {
          // Just try to get the active worksheet
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          worksheet.load('name');
          await context.sync();
          console.log('Excel connection verified, active worksheet:', worksheet.name);
          setIsExcelConnected(true);
        });
      } catch (error) {
        console.error('Excel connection test failed:', error);
        setIsExcelConnected(false);
      }
    };
    
    // Initial check
    checkExcelConnection();
    
    // Set up Office.onReady to detect when Office.js is fully initialized
    if (typeof Office !== 'undefined') {
      Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
          console.log('Office.js initialized in Excel');
          checkExcelConnection();
        } else {
          console.warn('Office.js initialized but not in Excel');
          setIsExcelConnected(false);
        }
      });
    }
    
    // Periodic check for connection status (less frequent to avoid performance issues)
    const interval = setInterval(checkExcelConnection, 10000);
    
    return () => clearInterval(interval);
  }, []);

  return (
    <ExcelContext.Provider 
      value={{
        activeCell, 
        selectedRange, 
        isExcelConnected,
        getWorksheetContext, 
        populateCell 
      }}
    >
      {children}
    </ExcelContext.Provider>
  );
};

export default ExcelContext;
