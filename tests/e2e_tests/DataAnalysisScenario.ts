/**
 * Data Analysis test scenario
 */
import { TestScenario } from "./TestScenario";

export const DataAnalysisScenario: TestScenario = {
  id: "data-analysis-001",
  name: "Data Analysis Test",
  description: "Tests data analysis features including sorting, filtering, and pivot tables",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "data-analysis-001-step-1",
      description: "Add dataset for analysis",
      query: "In Sheet1, create a dataset with headers in row 1: 'Date', 'Region', 'Product', 'Units', 'Revenue'. Then add the following data: \n1/15/2024, East, Laptop, 12, 10800 \n2/3/2024, West, Monitor, 15, 4500 \n1/22/2024, North, Laptop, 8, 7200 \n3/10/2024, East, Keyboard, 25, 1250 \n2/18/2024, South, Monitor, 10, 3000 \n3/5/2024, West, Keyboard, 30, 1500 \n1/30/2024, North, Mouse, 40, 1200 \n2/25/2024, South, Laptop, 5, 4500 \n3/15/2024, East, Mouse, 35, 1050",
      expectedOutcome: {
        cellValues: [
          // Headers
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "Date" },
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "Region" },
          { sheetName: "Sheet1", cellAddress: "C1", expectedValue: "Product" },
          { sheetName: "Sheet1", cellAddress: "D1", expectedValue: "Units" },
          { sheetName: "Sheet1", cellAddress: "E1", expectedValue: "Revenue" },
          
          // Sample data rows
          { sheetName: "Sheet1", cellAddress: "A2", expectedValue: "1/15/2024" },
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: "East" },
          { sheetName: "Sheet1", cellAddress: "C2", expectedValue: "Laptop" },
          { sheetName: "Sheet1", cellAddress: "D2", expectedValue: 12 },
          { sheetName: "Sheet1", cellAddress: "E2", expectedValue: 10800 },
          
          { sheetName: "Sheet1", cellAddress: "A10", expectedValue: "3/15/2024" },
          { sheetName: "Sheet1", cellAddress: "B10", expectedValue: "East" },
          { sheetName: "Sheet1", cellAddress: "C10", expectedValue: "Mouse" },
          { sheetName: "Sheet1", cellAddress: "D10", expectedValue: 35 },
          { sheetName: "Sheet1", cellAddress: "E10", expectedValue: 1050 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "data-analysis-001-step-2",
      description: "Sort the data by Revenue in descending order",
      query: "Sort the data in range A1:E10 by Revenue in descending order",
      expectedOutcome: {
        cellValues: [
          // Headers should remain the same
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "Date" },
          { sheetName: "Sheet1", cellAddress: "E1", expectedValue: "Revenue" },
          
          // First row should now be the highest revenue
          { sheetName: "Sheet1", cellAddress: "C2", expectedValue: "Laptop" },
          { sheetName: "Sheet1", cellAddress: "E2", expectedValue: 10800 }
        ],
        expectedOperations: ["sort_range"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "data-analysis-001-step-3",
      description: "Create a pivot table to analyze sales by region and product",
      query: "Create a pivot table in cell G1 using data from A1:E10, with Region as rows, Product as columns, and sum of Revenue as values",
      expectedOutcome: {
        expectedOperations: ["create_pivot_table"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "data-analysis-001-step-4",
      description: "Add data filtering",
      query: "Add filters to the headers in row 1",
      expectedOutcome: {
        expectedOperations: ["add_filter"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "data-analysis-001-step-5",
      description: "Clear the workbook for the next test",
      query: "Clear all data, filters, and pivot tables in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "G1", expectedValue: "" }
        ],
        expectedOperations: ["clear_contents", "remove_filter", "delete_pivot_table"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};
