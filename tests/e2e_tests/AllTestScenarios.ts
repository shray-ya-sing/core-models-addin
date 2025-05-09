/**
 * Test Scenario Model
 * Defines the structure for integration test scenarios
 */

/**
 * Expected outcome for a test step
 */
export interface ExpectedOutcome {
  // Expected cell values after operation
  cellValues?: Array<{
    sheetName: string;
    cellAddress: string;
    expectedValue: any;
  }>;
  
  // Expected cell formatting properties
  cellFormatting?: Array<{
    sheetName: string;
    cellAddress: string;
    properties: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontSize?: number;
      fontColor?: string;
      fillColor?: string;
      numberFormat?: string;
      horizontalAlignment?: string;
      verticalAlignment?: string;
      borders?: {
        top?: { style?: string; color?: string; weight?: string };
        bottom?: { style?: string; color?: string; weight?: string };
        left?: { style?: string; color?: string; weight?: string };
        right?: { style?: string; color?: string; weight?: string };
      };
      wrapText?: boolean;
      indentLevel?: number;
    };
  }>;
  
  // Expected chart formatting properties
  chartFormatting?: Array<{
    sheetName: string;
    chartName?: string;
    chartIndex?: number;
    properties: {
      chartType?: string;
      title?: {
        text?: string;
        visible?: boolean;
        fontSize?: number;
        bold?: boolean;
      };
      legend?: {
        position?: string;
        visible?: boolean;
      };
      hasDataLabels?: boolean;
      dataLabelsPosition?: string;
      seriesCount?: number;
      height?: number;
      width?: number;
    };
  }>;
  
  // Expected operations to be executed
  expectedOperations?: string[];
  
  // Expected query classification
  expectedQueryType?: string;
}

/**
 * Test step definition
 */
export interface TestStep {
  id: string;
  description: string;
  query: string;
  expectedOutcome: ExpectedOutcome;
}

/**
 * Test scenario definition
 */
export interface TestScenario {
  id: string;
  name: string;
  description: string;
  workbookPath: string;
  steps: TestStep[];
}

/**
 * Test step result
 */
export interface TestStepResult {
  stepId: string;
  query: string;
  success: boolean;
  message?: string;
  details?: {
    cellValueMatches?: Array<{
      address: string;
      expected: any;
      actual: any;
      match: boolean;
    }>;
    operationMatches?: Array<{
      operation: string;
      executed: boolean;
    }>;
    queryTypeMatch?: {
      expected: string;
      actual: string;
      match: boolean;
    };
  };
}

/**
 * Test result
 */
export interface TestResult {
  scenarioId: string;
  name: string;
  success: boolean;
  steps: TestStepResult[];
  startTime: number;
  endTime: number;
  duration: number;
}

/**
 * All test scenarios
 */

// Import all test scenarios
import { DataAnalysisScenario } from "./DataAnalysisScenario";
import { FormulaTestingScenario } from "./FormulaTestingScenario";
import { WorksheetManagementScenario } from "./WorksheetManagementScenario";
import { FinancialModelingScenario } from "./FinancialModelingScenario";

/**
 * Sample financial data test scenario
 */
export const SampleDataScenario: TestScenario = {
  id: "sample-data-001",
  name: "Add Sample Financial Data",
  description: "Tests adding sample financial data to a blank workbook",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "sample-data-001-step-1",
      description: "Add financial data with headers",
      query: "In Sheet1, add Q1-Q4 as headers in cells B1-E1, then add Revenue in cell A2 with values 120000, 145000, 160000, 190000 in cells B2-E2, Expenses in cell A3 with values 80000, 95000, 105000, 125000 in cells B3-E3, and Profit in cell A4 with values 40000, 50000, 55000, 65000 in cells B4-E4",
      expectedOutcome: {
        cellValues: [
          // Headers in row 1
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "Q1" },
          { sheetName: "Sheet1", cellAddress: "C1", expectedValue: "Q2" },
          { sheetName: "Sheet1", cellAddress: "D1", expectedValue: "Q3" },
          { sheetName: "Sheet1", cellAddress: "E1", expectedValue: "Q4" },
          
          // Revenue row
          { sheetName: "Sheet1", cellAddress: "A2", expectedValue: "Revenue" },
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: 120000 },
          { sheetName: "Sheet1", cellAddress: "C2", expectedValue: 145000 },
          { sheetName: "Sheet1", cellAddress: "D2", expectedValue: 160000 },
          { sheetName: "Sheet1", cellAddress: "E2", expectedValue: 190000 },
          
          // Expenses row
          { sheetName: "Sheet1", cellAddress: "A3", expectedValue: "Expenses" },
          { sheetName: "Sheet1", cellAddress: "B3", expectedValue: 80000 },
          { sheetName: "Sheet1", cellAddress: "C3", expectedValue: 95000 },
          { sheetName: "Sheet1", cellAddress: "D3", expectedValue: 105000 },
          { sheetName: "Sheet1", cellAddress: "E3", expectedValue: 125000 },
          
          // Profit row
          { sheetName: "Sheet1", cellAddress: "A4", expectedValue: "Profit" },
          { sheetName: "Sheet1", cellAddress: "B4", expectedValue: 40000 },
          { sheetName: "Sheet1", cellAddress: "C4", expectedValue: 50000 },
          { sheetName: "Sheet1", cellAddress: "D4", expectedValue: 55000 },
          { sheetName: "Sheet1", cellAddress: "E4", expectedValue: 65000 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "sample-data-001-step-2",
      description: "Clear the workbook for the next test",
      query: "Clear all data in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "C1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "D1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "E1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "A2", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: "" }
        ],
        expectedOperations: ["clear_contents"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};

/**
 * Chart creation test scenario
 */
export const ChartCreationScenario: TestScenario = {
  id: "chart-creation-001",
  name: "Chart Creation Test",
  description: "Tests creating various chart types from data",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "chart-creation-001-step-1",
      description: "Add data for charts",
      query: "In Sheet1, add 'Product' in cell H2, 'Q1' in cell I2, 'Q2' in cell J2, 'Q3' in cell K2, 'Q4' in cell L2. Then add 'Widgets' in cell H3 with values 250, 310, 380, 420 in cells I3-L3, 'Gadgets' in cell H4 with values 180, 250, 340, 360 in cells I4-L4, and 'Gizmos' in cell H5 with values 320, 290, 270, 400 in cells I5-L5.",
      expectedOutcome: {
        cellValues: [
          // Headers
          { sheetName: "Sheet1", cellAddress: "H2", expectedValue: "Product" },
          { sheetName: "Sheet1", cellAddress: "I2", expectedValue: "Q1" },
          { sheetName: "Sheet1", cellAddress: "J2", expectedValue: "Q2" },
          { sheetName: "Sheet1", cellAddress: "K2", expectedValue: "Q3" },
          { sheetName: "Sheet1", cellAddress: "L2", expectedValue: "Q4" },
          
          // Widgets row
          { sheetName: "Sheet1", cellAddress: "H3", expectedValue: "Widgets" },
          { sheetName: "Sheet1", cellAddress: "I3", expectedValue: 250 },
          { sheetName: "Sheet1", cellAddress: "J3", expectedValue: 310 },
          { sheetName: "Sheet1", cellAddress: "K3", expectedValue: 380 },
          { sheetName: "Sheet1", cellAddress: "L3", expectedValue: 420 },
          
          // Gadgets row
          { sheetName: "Sheet1", cellAddress: "H4", expectedValue: "Gadgets" },
          { sheetName: "Sheet1", cellAddress: "I4", expectedValue: 180 },
          { sheetName: "Sheet1", cellAddress: "J4", expectedValue: 250 },
          { sheetName: "Sheet1", cellAddress: "K4", expectedValue: 340 },
          { sheetName: "Sheet1", cellAddress: "L4", expectedValue: 360 },
          
          // Gizmos row
          { sheetName: "Sheet1", cellAddress: "H5", expectedValue: "Gizmos" },
          { sheetName: "Sheet1", cellAddress: "I5", expectedValue: 320 },
          { sheetName: "Sheet1", cellAddress: "J5", expectedValue: 290 },
          { sheetName: "Sheet1", cellAddress: "K5", expectedValue: 270 },
          { sheetName: "Sheet1", cellAddress: "L5", expectedValue: 400 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-creation-001-step-2",
      description: "Create a column chart",
      query: "Create a column chart using the data in range H2:L5, place it in cell H7, and give it the title 'Quarterly Product Sales'.",
      expectedOutcome: {
        expectedOperations: ["create_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-creation-001-step-3",
      description: "Create a line chart",
      query: "Create a line chart using the data in range H2:L5, place it in cell H20, and give it the title 'Product Sales Trends'.",
      expectedOutcome: {
        expectedOperations: ["create_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-creation-001-step-4",
      description: "Create a pie chart",
      query: "Create a pie chart showing the Q4 sales distribution by product (using data in range H2:H5 and L2:L5), place it in cell H33, and give it the title 'Q4 Sales Distribution'.",
      expectedOutcome: {
        expectedOperations: ["create_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-creation-001-step-5",
      description: "Clear the workbook for the next test",
      query: "Delete all charts and clear all data in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "H2", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "I2", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "J2", expectedValue: "" }
        ],
        expectedOperations: ["delete_chart", "clear_contents"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};

/**
 * Cell formatting test scenario
 */
export const CellFormattingScenario: TestScenario = {
  id: "cell-formatting-001",
  name: "Cell Formatting Test",
  description: "Tests applying various cell formatting options to a blank workbook",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "cell-formatting-001-step-1",
      description: "Add data for formatting",
      query: "In Sheet1, add 'Sales Report' as a title in cell B2, add 'Region' in cell B4, 'Q1' in cell C4, 'Q2' in cell D4, 'Q3' in cell E4, 'Q4' in cell F4, then add 'North' in cell B5 with values 45000, 52000, 61000, 70000 in cells C5-F5, 'South' in cell B6 with values 38000, 44000, 53000, 65000 in cells C6-F6, and 'Total' in cell B7 with formulas to sum each quarter in cells C7-F7",
      expectedOutcome: {
        cellValues: [
          // Title
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: "Sales Report" },
          
          // Headers
          { sheetName: "Sheet1", cellAddress: "B4", expectedValue: "Region" },
          { sheetName: "Sheet1", cellAddress: "C4", expectedValue: "Q1" },
          { sheetName: "Sheet1", cellAddress: "D4", expectedValue: "Q2" },
          { sheetName: "Sheet1", cellAddress: "E4", expectedValue: "Q3" },
          { sheetName: "Sheet1", cellAddress: "F4", expectedValue: "Q4" },
          
          // North region
          { sheetName: "Sheet1", cellAddress: "B5", expectedValue: "North" },
          { sheetName: "Sheet1", cellAddress: "C5", expectedValue: 45000 },
          { sheetName: "Sheet1", cellAddress: "D5", expectedValue: 52000 },
          { sheetName: "Sheet1", cellAddress: "E5", expectedValue: 61000 },
          { sheetName: "Sheet1", cellAddress: "F5", expectedValue: 70000 },
          
          // South region
          { sheetName: "Sheet1", cellAddress: "B6", expectedValue: "South" },
          { sheetName: "Sheet1", cellAddress: "C6", expectedValue: 38000 },
          { sheetName: "Sheet1", cellAddress: "D6", expectedValue: 44000 },
          { sheetName: "Sheet1", cellAddress: "E6", expectedValue: 53000 },
          { sheetName: "Sheet1", cellAddress: "F6", expectedValue: 65000 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "cell-formatting-001-step-2",
      description: "Apply formatting to the title and headers",
      query: "Format the title 'Sales Report' in cell B2 with bold, font size 16, and center it across columns B through F. Format the headers in row 4 with bold, background color light blue, and center alignment.",
      expectedOutcome: {
        expectedOperations: ["set_font", "set_fill", "set_alignment", "merge_cells"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "cell-formatting-001-step-3",
      description: "Apply number formatting to the data cells",
      query: "Format all values in cells C5:F6 as currency with 0 decimal places. Add a bottom border to the total row (row 7). Apply conditional formatting to highlight values above 60000 in green and below 40000 in light red.",
      expectedOutcome: {
        expectedOperations: ["set_number_format", "set_border", "set_conditional_format"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "cell-formatting-001-step-4",
      description: "Clear the workbook for the next test",
      query: "Clear all data and formatting in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B4", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "C4", expectedValue: "" }
        ],
        expectedOperations: ["clear_contents", "clear_formats"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};

/**
 * Chart formatting test scenario
 */
export const ChartFormattingScenario: TestScenario = {
  id: "chart-formatting-001",
  name: "Chart Formatting Test",
  description: "Tests formatting options for charts",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "chart-formatting-001-step-1",
      description: "Add data for chart formatting",
      query: "In Sheet1, add 'Month' in cell N2, 'Revenue' in cell O2, 'Expenses' in cell P2, 'Profit' in cell Q2. Then add the following data: January in N3 with values 85000, 52000, 33000 in O3-Q3, February in N4 with values 92000, 56000, 36000 in O4-Q4, March in N5 with values 103000, 61000, 42000 in O5-Q5, April in N6 with values 116000, 68000, 48000 in O6-Q6, May in N7 with values 124000, 71000, 53000 in O7-Q7, June in N8 with values 138000, 76000, 62000 in O8-Q8.",
      expectedOutcome: {
        cellValues: [
          // Headers
          { sheetName: "Sheet1", cellAddress: "N2", expectedValue: "Month" },
          { sheetName: "Sheet1", cellAddress: "O2", expectedValue: "Revenue" },
          { sheetName: "Sheet1", cellAddress: "P2", expectedValue: "Expenses" },
          { sheetName: "Sheet1", cellAddress: "Q2", expectedValue: "Profit" },
          
          // January row
          { sheetName: "Sheet1", cellAddress: "N3", expectedValue: "January" },
          { sheetName: "Sheet1", cellAddress: "O3", expectedValue: 85000 },
          { sheetName: "Sheet1", cellAddress: "P3", expectedValue: 52000 },
          { sheetName: "Sheet1", cellAddress: "Q3", expectedValue: 33000 },
          
          // February row
          { sheetName: "Sheet1", cellAddress: "N4", expectedValue: "February" },
          { sheetName: "Sheet1", cellAddress: "O4", expectedValue: 92000 },
          { sheetName: "Sheet1", cellAddress: "P4", expectedValue: 56000 },
          { sheetName: "Sheet1", cellAddress: "Q4", expectedValue: 36000 },
          
          // March row
          { sheetName: "Sheet1", cellAddress: "N5", expectedValue: "March" },
          { sheetName: "Sheet1", cellAddress: "O5", expectedValue: 103000 },
          { sheetName: "Sheet1", cellAddress: "P5", expectedValue: 61000 },
          { sheetName: "Sheet1", cellAddress: "Q5", expectedValue: 42000 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-formatting-001-step-2",
      description: "Create a basic chart for formatting",
      query: "Create a line chart using data in range N2:Q8, place it in cell N10, and give it the title 'Financial Performance H1 2025'.",
      expectedOutcome: {
        expectedOperations: ["create_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-formatting-001-step-3",
      description: "Format chart title and legend",
      query: "Format the chart titled 'Financial Performance H1 2025' by making the title bold, font size 14, and dark blue color. Position the chart legend at the bottom and make it horizontal.",
      expectedOutcome: {
        expectedOperations: ["format_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-formatting-001-step-4",
      description: "Format chart axes and data series",
      query: "For the chart titled 'Financial Performance H1 2025', format the vertical axis to show currency values with no decimal places and add a title 'Amount (USD)'. Format the Revenue series in blue, Expenses series in red, and Profit series in green. Add data labels to the Profit series only.",
      expectedOutcome: {
        expectedOperations: ["format_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-formatting-001-step-5",
      description: "Add and format a trendline",
      query: "For the chart titled 'Financial Performance H1 2025', add a linear trendline to the Revenue series, format it with a dashed line style, and display the equation on the chart.",
      expectedOutcome: {
        expectedOperations: ["format_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "chart-formatting-001-step-6",
      description: "Clear the workbook for the next test",
      query: "Delete all charts and clear all data in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "N2", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "O2", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "P2", expectedValue: "" }
        ],
        expectedOperations: ["delete_chart", "clear_contents"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};

/**
 * Array of all test scenarios
 */
export const AllTestScenarios = [
  SampleDataScenario,
  ChartCreationScenario,
  CellFormattingScenario,
  ChartFormattingScenario,
  DataAnalysisScenario,
  FormulaTestingScenario,
  WorksheetManagementScenario,
  FinancialModelingScenario
];
