/**
 * Worksheet Management Test Scenario
 * Tests Excel worksheet manipulation capabilities
 */
import { TestScenario } from "./TestScenario";

export const WorksheetManagementScenario: TestScenario = {
  id: "worksheet-mgmt-001",
  name: "Worksheet Management Test",
  description: "Tests Excel worksheet manipulation capabilities including adding, renaming, and formatting worksheets",
  workbookPath: "./test-workbooks/blank.xlsx",
  steps: [
    {
      id: "worksheet-mgmt-001-step-1",
      description: "Add new worksheets",
      query: "Add three new worksheets named 'Income Statement', 'Balance Sheet', and 'Cash Flow'",
      expectedOutcome: {
        expectedOperations: ["add_worksheet"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "worksheet-mgmt-001-step-2",
      description: "Add basic data to each worksheet",
      query: "In the 'Income Statement' worksheet, add 'Revenue' in cell A1 and '100000' in cell B1. In the 'Balance Sheet' worksheet, add 'Assets' in cell A1 and '250000' in cell B1. In the 'Cash Flow' worksheet, add 'Operating Cash Flow' in cell A1 and '75000' in cell B1.",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Income Statement", cellAddress: "A1", expectedValue: "Revenue" },
          { sheetName: "Income Statement", cellAddress: "B1", expectedValue: 100000 },
          { sheetName: "Balance Sheet", cellAddress: "A1", expectedValue: "Assets" },
          { sheetName: "Balance Sheet", cellAddress: "B1", expectedValue: 250000 },
          { sheetName: "Cash Flow", cellAddress: "A1", expectedValue: "Operating Cash Flow" },
          { sheetName: "Cash Flow", cellAddress: "B1", expectedValue: 75000 }
        ],
        expectedOperations: ["set_values"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "worksheet-mgmt-001-step-3",
      description: "Format worksheet tabs with colors",
      query: "Set the tab color of 'Income Statement' to green, 'Balance Sheet' to blue, and 'Cash Flow' to orange",
      expectedOutcome: {
        expectedOperations: ["set_tab_color"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "worksheet-mgmt-001-step-4",
      description: "Reorder worksheets",
      query: "Move the 'Balance Sheet' worksheet to be the first worksheet in the workbook",
      expectedOutcome: {
        expectedOperations: ["move_worksheet"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "worksheet-mgmt-001-step-5",
      description: "Add worksheet protection",
      query: "Protect the 'Income Statement' worksheet with password 'test123' to prevent users from modifying cells",
      expectedOutcome: {
        expectedOperations: ["protect_worksheet"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "worksheet-mgmt-001-step-6",
      description: "Set worksheet view options",
      query: "In the 'Balance Sheet' worksheet, hide gridlines and set zoom level to 85%",
      expectedOutcome: {
        expectedOperations: ["set_gridlines", "set_zoom"],
        expectedQueryType: "workbook_command"
      }
    },
    {
      id: "worksheet-mgmt-001-step-7",
      description: "Clean up for next test",
      query: "Delete all worksheets except Sheet1, and clear all data in Sheet1",
      expectedOutcome: {
        cellValues: [
          { sheetName: "Sheet1", cellAddress: "A1", expectedValue: "" }
        ],
        expectedOperations: ["delete_worksheet", "clear_contents"],
        expectedQueryType: "workbook_command"
      }
    }
  ]
};
