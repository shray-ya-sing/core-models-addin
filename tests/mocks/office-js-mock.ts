/**
 * Mock implementation of Office.js API for unit testing
 */

// Mock Excel API components
export const mockExcelApi = {
  // Mock context
  context: {
    sync: jest.fn().mockImplementation(async callback => {
      if (callback) {
        await callback();
      }
      return Promise.resolve();
    }),
    trackedObjects: {
      add: jest.fn(),
      remove: jest.fn()
    }
  },

  // Mock workbook
  workbook: {
    context: {},
    worksheets: {
      getActiveWorksheet: jest.fn(),
      getItem: jest.fn(),
      getItemOrNullObject: jest.fn(),
      items: [],
      load: jest.fn(),
      _RegisteredObjectArray: true
    },
    getSelectedRange: jest.fn(),
    load: jest.fn()
  },

  // Mock worksheet
  worksheet: {
    name: "Sheet1",
    position: 0,
    visibility: "Visible",
    getRange: jest.fn(),
    getUsedRange: jest.fn().mockReturnValue({
      address: "A1:D10",
      columnCount: 4,
      rowCount: 10,
      values: [],
      formulas: []
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
  },

  // Mock range
  range: {
    address: "Sheet1!A1:B2",
    values: [["Value1", "Value2"], ["Value3", "Value4"]],
    formulas: [["=SUM(A1)", "=A1*2"], ["", ""]],
    format: {
      fill: {
        color: "#FFFFFF"
      },
      font: {
        bold: false,
        italic: false,
        name: "Calibri",
        size: 11,
        color: "#000000"
      },
      borders: {},
      horizontalAlignment: "General",
      verticalAlignment: "Bottom"
    },
    load: jest.fn(),
    getRow: jest.fn(),
    getColumn: jest.fn(),
    getIntersection: jest.fn(),
    getIntersectionOrNullObject: jest.fn()
  },

  // Mock table
  table: {
    name: "Table1",
    getRange: jest.fn(),
    columns: {
      getItemOrNullObject: jest.fn(),
      items: [],
      _RegisteredObjectArray: true
    },
    rows: {
      getItemOrNullObject: jest.fn(),
      items: [],
      _RegisteredObjectArray: true
    },
    load: jest.fn()
  },

  // Mock chart
  chart: {
    name: "Chart1",
    chartType: "ColumnClustered",
    load: jest.fn()
  }
};

// Create mock Office namespace
export const mockOffice = {
  context: {
    requirements: {
      isSetSupported: jest.fn().mockReturnValue(true)
    }
  },
  Excel: {
    run: jest.fn().mockImplementation(async callback => {
      return await callback(mockExcelApi.context);
    })
  }
};

// Set up global Office mock
global.Office = mockOffice;
