/**
 * Financial Modeling Test Scenario
 * Tests Excel capabilities for financial modeling, including tax calculations and yield curve analysis
 */
import { TestScenario } from "./TestScenario";

export const FinancialModelingScenario: TestScenario = {
  id: "financial-modeling-001",
  name: "Financial Modeling Test",
  description: "Tests Excel capabilities for financial modeling, including tax calculations and yield curve analysis",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "financial-modeling-001-step-1",
      description: "Add income statement with book vs tax differences",
      query: "In Sheet1, create an income statement with the following structure: In cell A1 add 'Income Statement', in A3 add 'Revenue', in A4 add 'COGS', in A5 add 'Gross Profit', in A6 add 'Operating Expenses', in A7 add 'Depreciation', in A8 add 'Operating Income', in A9 add 'Interest Expense', in A10 add 'Income Before Tax', in A11 add 'Tax Expense', in A12 add 'Net Income'. Add column headers 'Book' in B2 and 'Tax' in C2. Add the following values in the Book column: 1000000, 600000, formula for B3-B4, 150000, 50000, formula for B5-B6-B7, 30000, formula for B8-B9, formula for B10*0.25, formula for B10-B11. Add the following values in the Tax column: 1000000, 600000, formula for C3-C4, 150000, 80000, formula for C5-C6-C7, 30000, formula for C8-C9, formula for C10*0.25, formula for C10-C11.",
      expectedOutcome: {
        cellValues: [
          // Headers
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "Income Statement" },
          { sheetName: "Sheet1", cellAddress: "B2", expectedValue: "Book" },
          { sheetName: "Sheet1", cellAddress: "C2", expectedValue: "Tax" },
          
          // Revenue
          { sheetName: "Sheet1", cellAddress: "A3", expectedValue: "Revenue" },
          { sheetName: "Sheet1", cellAddress: "B3", expectedValue: 1000000 },
          { sheetName: "Sheet1", cellAddress: "C3", expectedValue: 1000000 },
          
          // COGS
          { sheetName: "Sheet1", cellAddress: "A4", expectedValue: "COGS" },
          { sheetName: "Sheet1", cellAddress: "B4", expectedValue: 600000 },
          { sheetName: "Sheet1", cellAddress: "C4", expectedValue: 600000 },
          
          // Gross Profit (calculated)
          { sheetName: "Sheet1", cellAddress: "A5", expectedValue: "Gross Profit" },
          { sheetName: "Sheet1", cellAddress: "B5", expectedValue: 400000 }, // B3-B4
          { sheetName: "Sheet1", cellAddress: "C5", expectedValue: 400000 }, // C3-C4
          
          // Operating Expenses
          { sheetName: "Sheet1", cellAddress: "A6", expectedValue: "Operating Expenses" },
          { sheetName: "Sheet1", cellAddress: "B6", expectedValue: 150000 },
          { sheetName: "Sheet1", cellAddress: "C6", expectedValue: 150000 },
          
          // Depreciation (different between book and tax)
          { sheetName: "Sheet1", cellAddress: "A7", expectedValue: "Depreciation" },
          { sheetName: "Sheet1", cellAddress: "B7", expectedValue: 50000 },
          { sheetName: "Sheet1", cellAddress: "C7", expectedValue: 80000 }, // Accelerated depreciation for tax
          
          // Operating Income (calculated)
          { sheetName: "Sheet1", cellAddress: "A8", expectedValue: "Operating Income" },
          { sheetName: "Sheet1", cellAddress: "B8", expectedValue: 200000 }, // B5-B6-B7
          { sheetName: "Sheet1", cellAddress: "C8", expectedValue: 170000 }, // C5-C6-C7
          
          // Interest Expense
          { sheetName: "Sheet1", cellAddress: "A9", expectedValue: "Interest Expense" },
          { sheetName: "Sheet1", cellAddress: "B9", expectedValue: 30000 },
          { sheetName: "Sheet1", cellAddress: "C9", expectedValue: 30000 },
          
          // Income Before Tax (calculated)
          { sheetName: "Sheet1", cellAddress: "A10", expectedValue: "Income Before Tax" },
          { sheetName: "Sheet1", cellAddress: "B10", expectedValue: 170000 }, // B8-B9
          { sheetName: "Sheet1", cellAddress: "C10", expectedValue: 140000 }, // C8-C9
          
          // Tax Expense (calculated)
          { sheetName: "Sheet1", cellAddress: "A11", expectedValue: "Tax Expense" },
          { sheetName: "Sheet1", cellAddress: "B11", expectedValue: 42500 }, // B10*0.25
          { sheetName: "Sheet1", cellAddress: "C11", expectedValue: 35000 }, // C10*0.25
          
          // Net Income (calculated)
          { sheetName: "Sheet1", cellAddress: "A12", expectedValue: "Net Income" },
          { sheetName: "Sheet1", cellAddress: "B12", expectedValue: 127500 }, // B10-B11
          { sheetName: "Sheet1", cellAddress: "C12", expectedValue: 105000 }  // C10-C11
        ],
        expectedOperations: ["set_values", "set_formula"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "financial-modeling-001-step-2",
      description: "Add deferred tax calculation",
      query: "In cell E1, add 'Deferred Tax Analysis'. In cell E3, add 'Temporary Differences', in cell E4 add 'Book Income Before Tax', in cell E5 add 'Taxable Income', in cell E6 add 'Difference', in cell E7 add 'Tax Rate', in cell E8 add 'Deferred Tax Asset/(Liability)'. In column F, add the following: in F4 reference cell B10, in F5 reference cell C10, in F6 add a formula to calculate F4-F5, in F7 add 25%, in F8 add a formula to calculate F6*F7.",
      expectedOutcome: {
        cellValues: [
          // Headers
          { sheetName: "Sheet1", cellAddress: "E1", expectedValue: "Deferred Tax Analysis" },
          { sheetName: "Sheet1", cellAddress: "E3", expectedValue: "Temporary Differences" },
          
          // Book Income
          { sheetName: "Sheet1", cellAddress: "E4", expectedValue: "Book Income Before Tax" },
          { sheetName: "Sheet1", cellAddress: "F4", expectedValue: 170000 }, // Reference to B10
          
          // Taxable Income
          { sheetName: "Sheet1", cellAddress: "E5", expectedValue: "Taxable Income" },
          { sheetName: "Sheet1", cellAddress: "F5", expectedValue: 140000 }, // Reference to C10
          
          // Difference
          { sheetName: "Sheet1", cellAddress: "E6", expectedValue: "Difference" },
          { sheetName: "Sheet1", cellAddress: "F6", expectedValue: 30000 }, // F4-F5
          
          // Tax Rate
          { sheetName: "Sheet1", cellAddress: "E7", expectedValue: "Tax Rate" },
          { sheetName: "Sheet1", cellAddress: "F7", expectedValue: 0.25 }, // 25%
          
          // Deferred Tax Asset
          { sheetName: "Sheet1", cellAddress: "E8", expectedValue: "Deferred Tax Asset/(Liability)" },
          { sheetName: "Sheet1", cellAddress: "F8", expectedValue: 7500 } // F6*F7
        ],
        expectedOperations: ["set_values", "set_formula"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "financial-modeling-001-step-3",
      description: "Add yield curve analysis",
      query: "In a new worksheet named 'Yield Curve', add the following: In cell A1 add 'Yield Curve Analysis', in cell A3 add 'Maturity (Years)', in cells B3:F3 add 1, 2, 5, 10, 30. In cell A4 add 'Yield (%)', in cells B4:F4 add 4.5, 4.7, 5.0, 5.2, 5.5. In cell A6 add 'Nelson-Siegel Parameters', in cell A7 add 'Beta0 (long-term)', in cell A8 add 'Beta1 (short-term)', in cell A9 add 'Beta2 (medium-term)', in cell A10 add 'Tau (decay factor)', in cells B7:B10 add 5.5, -1.0, 0.5, 2.0. In cell A12 add 'Fitted Yields', in cells A13:A17 add the same maturities as in B3:F3. In cells B13:B17 add formulas to calculate the Nelson-Siegel yield for each maturity using parameters in B7:B10.",
      expectedOutcome: {
        cellValues: [
          // Headers
          { sheetName: "Yield Curve", cellAddress: "A1", expectedValue: "Yield Curve Analysis" },
          { sheetName: "Yield Curve", cellAddress: "A3", expectedValue: "Maturity (Years)" },
          
          // Maturities
          { sheetName: "Yield Curve", cellAddress: "B3", expectedValue: 1 },
          { sheetName: "Yield Curve", cellAddress: "C3", expectedValue: 2 },
          { sheetName: "Yield Curve", cellAddress: "D3", expectedValue: 5 },
          { sheetName: "Yield Curve", cellAddress: "E3", expectedValue: 10 },
          { sheetName: "Yield Curve", cellAddress: "F3", expectedValue: 30 },
          
          // Yields
          { sheetName: "Yield Curve", cellAddress: "A4", expectedValue: "Yield (%)" },
          { sheetName: "Yield Curve", cellAddress: "B4", expectedValue: 4.5 },
          { sheetName: "Yield Curve", cellAddress: "C4", expectedValue: 4.7 },
          { sheetName: "Yield Curve", cellAddress: "D4", expectedValue: 5.0 },
          { sheetName: "Yield Curve", cellAddress: "E4", expectedValue: 5.2 },
          { sheetName: "Yield Curve", cellAddress: "F4", expectedValue: 5.5 },
          
          // Nelson-Siegel Parameters
          { sheetName: "Yield Curve", cellAddress: "A6", expectedValue: "Nelson-Siegel Parameters" },
          { sheetName: "Yield Curve", cellAddress: "A7", expectedValue: "Beta0 (long-term)" },
          { sheetName: "Yield Curve", cellAddress: "B7", expectedValue: 5.5 },
          { sheetName: "Yield Curve", cellAddress: "A8", expectedValue: "Beta1 (short-term)" },
          { sheetName: "Yield Curve", cellAddress: "B8", expectedValue: -1.0 },
          { sheetName: "Yield Curve", cellAddress: "A9", expectedValue: "Beta2 (medium-term)" },
          { sheetName: "Yield Curve", cellAddress: "B9", expectedValue: 0.5 },
          { sheetName: "Yield Curve", cellAddress: "A10", expectedValue: "Tau (decay factor)" },
          { sheetName: "Yield Curve", cellAddress: "B10", expectedValue: 2.0 }
        ],
        expectedOperations: ["add_worksheet", "set_values", "set_formula"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "financial-modeling-001-step-4",
      description: "Create a yield curve chart",
      query: "Create a line chart in the Yield Curve worksheet using the data in cells A3:F4, place it in cell A20, and give it the title 'Yield Curve'. Add a second series to the chart using the fitted yields from cells A12:B17.",
      expectedOutcome: {
        expectedOperations: ["create_chart"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "financial-modeling-001-step-5",
      description: "Apply financial formatting",
      query: "Format all monetary values in the Sheet1 worksheet as Currency with 0 decimal places. Format the tax rate as Percentage with 2 decimal places. Add a thick bottom border to the Net Income row. Make the Income Statement title bold and font size 14.",
      expectedOutcome: {
        expectedOperations: ["set_number_format", "set_border", "set_font"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "financial-modeling-001-step-6",
      description: "Clear the workbook for the next test",
      query: "Delete all worksheets except Sheet1, and clear all data and formatting in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "" },
          { sheetName: "Sheet1", cellAddress: "B1", expectedValue: "" }
        ],
        expectedOperations: ["delete_worksheet", "clear_contents", "clear_formats"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};
