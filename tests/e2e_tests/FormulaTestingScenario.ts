/**
 * Formula Testing Scenario
 * Tests various Excel formula capabilities
 */
import { TestScenario } from "./TestScenario";

export const FormulaTestingScenario: TestScenario = {
  id: "formula-testing-001",
  name: "Formula Testing",
  description: "Tests various Excel formula capabilities including basic calculations, lookups, and conditional formulas",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "formula-testing-001-step-1",
      description: "Add base data for formula testing",
      query: "In Sheet1, add 'Values' as header in cell A1, then add numbers 10, 20, 30, 40, 50 in cells A2:A6. In cell C1 add 'Products', in cell D1 add 'Price', then add the following: 'Laptop', 1200 in C2:D2, 'Monitor', 300 in C3:D3, 'Keyboard', 50 in C4:D4, 'Mouse', 30 in C5:D5, 'Headset', 100 in C6:D6.",
      expectedOutcome: {
        cellValues: [
          // Values column
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "Values" },
          { sheetName: "Sheet1", cellAddress: "A2", expectedValue: 10 },
          { sheetName: "Sheet1", cellAddress: "A3", expectedValue: 20 },
          { sheetName: "Sheet1", cellAddress: "A4", expectedValue: 30 },
          { sheetName: "Sheet1", cellAddress: "A5", expectedValue: 40 },
          { sheetName: "Sheet1", cellAddress: "A6", expectedValue: 50 },
          
          // Products and prices
          { sheetName: "Sheet1", cellAddress: "C1", expectedValue: "Products" },
          { sheetName: "Sheet1", cellAddress: "D1", expectedValue: "Price" },
          { sheetName: "Sheet1", cellAddress: "C2", expectedValue: "Laptop" },
          { sheetName: "Sheet1", cellAddress: "D2", expectedValue: 1200 },
          { sheetName: "Sheet1", cellAddress: "C3", expectedValue: "Monitor" },
          { sheetName: "Sheet1", cellAddress: "D3", expectedValue: 300 },
          { sheetName: "Sheet1", cellAddress: "C6", expectedValue: "Headset" },
          { sheetName: "Sheet1", cellAddress: "D6", expectedValue: 100 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "formula-testing-001-step-2",
      description: "Add basic arithmetic formulas",
      query: "In cell B1, add 'Formula Results'. In cell B2, add a formula to calculate the sum of values in A2:A6. In cell B3, add a formula to calculate the average of values in A2:A6. In cell B4, add a formula to find the maximum value in A2:A6. In cell B5, add a formula to find the minimum value in A2:A6.",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "Formula Results" },
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: 150 }, // Sum of 10+20+30+40+50
          { sheetName: "Sheet1", cellAddress: "B3", expectedValue: 30 },  // Average of values
          { sheetName: "Sheet1", cellAddress: "B4", expectedValue: 50 },  // Maximum value
          { sheetName: "Sheet1", cellAddress: "B5", expectedValue: 10 }   // Minimum value
        ],
        expectedOperations: ["set_values", "set_formula"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "formula-testing-001-step-3",
      description: "Add lookup formulas",
      query: "In cell F1, add 'Lookup Tests'. In cell F2, add 'Product' and in cell G2 add 'Laptop'. In cell F3, add 'Price' and in cell G3 add a VLOOKUP formula to find the price of the product in G2 by looking it up in the range C2:D6.",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "F1", expectedValue: "Lookup Tests" },
          { sheetName: "Sheet1", cellAddress: "F2", expectedValue: "Product" },
          { sheetName: "Sheet1", cellAddress: "G2", expectedValue: "Laptop" },
          { sheetName: "Sheet1", cellAddress: "F3", expectedValue: "Price" },
          { sheetName: "Sheet1", cellAddress: "G3", expectedValue: 1200 } // Result of VLOOKUP
        ],
        expectedOperations: ["set_values", "set_formula"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "formula-testing-001-step-4",
      description: "Add conditional formulas",
      query: "In cell F5, add 'Conditional Tests'. In cell F6, add 'Value' and in cell G6 add 25. In cell F7, add 'Category' and in cell G7 add an IF formula that returns 'High' if G6 is greater than 50, 'Medium' if G6 is greater than 20, and 'Low' otherwise.",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "F5", expectedValue: "Conditional Tests" },
          { sheetName: "Sheet1", cellAddress: "F6", expectedValue: "Value" },
          { sheetName: "Sheet1", cellAddress: "G6", expectedValue: 25 },
          { sheetName: "Sheet1", cellAddress: "F7", expectedValue: "Category" },
          { sheetName: "Sheet1", cellAddress: "G7", expectedValue: "Medium" } // Result of IF formula
        ],
        expectedOperations: ["set_values", "set_formula"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "formula-testing-001-step-5",
      description: "Clear the workbook for the next test",
      query: "Clear all data and formulas in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "F1", expectedValue: "" }
        ],
        expectedOperations: ["clear_contents"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};
