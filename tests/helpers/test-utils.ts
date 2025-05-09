/**
 * Test utility functions to support unit testing
 */

/**
 * Create a mock worksheet with the given name and data
 * @param name Worksheet name
 * @param data Optional 2D array of cell values
 * @returns Mock worksheet object
 */
export function createMockWorksheet(name: string, data?: any[][]) {
  return {
    name,
    position: 0,
    visibility: "Visible",
    getRange: jest.fn().mockImplementation((address: string) => {
      return {
        address: `${name}!${address}`,
        values: data || [["Value1", "Value2"], ["Value3", "Value4"]],
        formulas: [["=SUM(A1)", "=A1*2"], ["", ""]],
        format: {
          fill: { color: "#FFFFFF" },
          font: { bold: false, italic: false, name: "Calibri", size: 11, color: "#000000" },
          horizontalAlignment: "General",
          verticalAlignment: "Bottom"
        },
        load: jest.fn()
      };
    }),
    getUsedRange: jest.fn().mockReturnValue({
      address: "A1:D10",
      columnCount: 4,
      rowCount: 10,
      values: data || Array(10).fill(Array(4).fill("")),
      formulas: Array(10).fill(Array(4).fill("")),
      load: jest.fn()
    }),
    tables: {
      getItemOrNullObject: jest.fn(),
      items: [],
      _RegisteredObjectArray: true
    },
    charts: {
      getItemOrNullObject: jest.fn(),
      items: [],
      _RegisteredObjectArray: true
    },
    load: jest.fn()
  };
}

/**
 * Create a mock workbook with the given worksheets
 * @param worksheetNames Array of worksheet names to create
 * @returns Mock workbook object
 */
export function createMockWorkbook(worksheetNames: string[]) {
  const worksheetItems = worksheetNames.map((name, index) => 
    createMockWorksheet(name));
  
  return {
    context: {},
    worksheets: {
      getActiveWorksheet: jest.fn().mockReturnValue(worksheetItems[0]),
      getItem: jest.fn().mockImplementation((name: string) => {
        const found = worksheetItems.find(ws => ws.name === name);
        if (!found) throw new Error(`Worksheet ${name} not found`);
        return found;
      }),
      getItemOrNullObject: jest.fn().mockImplementation((name: string) => {
        return worksheetItems.find(ws => ws.name === name) || null;
      }),
      items: worksheetItems,
      load: jest.fn(),
      _RegisteredObjectArray: true
    },
    getSelectedRange: jest.fn(),
    load: jest.fn()
  };
}

/**
 * Wait for the next event loop tick
 * Useful for testing async code
 */
export async function flushPromises() {
  return new Promise(resolve => setTimeout(resolve, 0));
}
