# Excel Add-in Testing Guide

## Introduction

Thank you for helping test the Excel Financial Models Add-in! This guide will walk you through what the add-in can currently do and provide a structured approach to testing its capabilities.

## Current Capabilities

The add-in is currently in development with the following status:

- ✅ **Workbook Analysis**: The add-in can analyze and explain Excel workbooks, worksheets, and cell ranges
- ✅ **Question Answering**: The add-in can answer questions about financial models and their components
- ✅ **Formula Explanation**: The add-in can explain complex formulas and their purpose
- ❌ **Knowledge Base Integration**: External knowledge integration is not yet functional
- ✅ **Command Execution**: The add-in can now make modifications to the workbook with 23 different operation types

**Important**: The add-in has been enhanced with command execution capabilities. Please test both explanation features and the new command execution operations listed below.

## How to Test

1. **Open a financial model** in Excel (ideally a model with multiple sheets, formulas, and financial concepts)
2. **Start the add-in** following the installation guide
3. **Ask questions** from the test cases below
4. **Record results** by marking each test as:
   - ✅ Success: The add-in provided a helpful, accurate response
   - ⚠️ Partial: The add-in understood but gave incomplete information
   - ❌ Failure: The add-in misunderstood or gave incorrect information
5. **Provide feedback** on what worked well or could be improved

## Testing Categories

The testing is organized into four levels:

1. **Workbook-level**: Questions about the entire financial model
2. **Worksheet-level**: Questions about specific worksheets
3. **Cell Range-level**: Questions about specific cells or ranges
4. **Command Execution**: Testing the add-in's ability to modify the workbook

## Test Cases

Some sample questions that might be useful for testing the different scenarios are suggested below. For command execution, we've provided a comprehensive set of tests organized by operation type.

### Workbook-Level Questions

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | What type of financial model is this? | | |
| 2 | Explain the overall structure of this workbook. | | |
| 3 | What are the key assumptions in this model? | | |
| 4 | How are the different sheets in this workbook connected? | | |
| 5 | What are the main outputs or KPIs in this model? | | |
| 6 | Explain the cash flow projections in this model. | | |
| 7 | What valuation methods are used in this model? | | |
| 8 | How are growth rates applied throughout this model? | | |
| 9 | What time period does this financial model cover? | | |
| 10 | Identify the inputs and outputs of this financial model. | | |
| 11 | What financial metrics are calculated in this workbook? | | |
| 12 | How is debt modeled in this workbook? | | |
| 13 | Explain how revenue forecasting works in this model. | | |
| 14 | What are the key drivers of profitability in this model? | | |
| 15 | How are taxes calculated in this financial model? | | |
| 16 | What sensitivity analyses are included in this model? | | |
| 17 | Explain the capital structure assumptions in this model. | | |
| 18 | How are working capital requirements calculated? | | |
| 19 | What discount rates are used in this model and where? | | |
| 20 | Explain the relationship between the income statement, balance sheet, and cash flow statement in this model. | | |

### Worksheet-Level Questions

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | What is the purpose of the [Sheet Name] worksheet? | | |
| 2 | Explain the calculations on the [Sheet Name] sheet. | | |
| 3 | What are the key inputs on the [Sheet Name] worksheet? | | |
| 4 | How does the [Sheet Name] sheet connect to other sheets? | | |
| 5 | What formulas are used most frequently on the [Sheet Name] sheet? | | |
| 6 | Explain the structure of the [Sheet Name] worksheet. | | |
| 7 | What are the main sections of the [Sheet Name] worksheet? | | |
| 8 | How are totals calculated on the [Sheet Name] sheet? | | |
| 9 | What assumptions are made on the [Sheet Name] worksheet? | | |
| 10 | Explain the color coding used on the [Sheet Name] sheet. | | |
| 11 | What time periods are covered in the [Sheet Name] worksheet? | | |
| 12 | How are growth rates applied in the [Sheet Name] sheet? | | |
| 13 | What are the key outputs from the [Sheet Name] worksheet? | | |
| 14 | Explain the tables in the [Sheet Name] worksheet. | | |
| 15 | What charts or graphs are on the [Sheet Name] sheet and what do they show? | | |
| 16 | How are percentages calculated on the [Sheet Name] worksheet? | | |
| 17 | What conditional formatting is used on the [Sheet Name] sheet? | | |
| 18 | Explain how data flows through the [Sheet Name] worksheet. | | |
| 19 | What are the most important cells on the [Sheet Name] sheet? | | |
| 20 | How does the [Sheet Name] worksheet contribute to the overall model? | | |

### Cell Range-Level Questions

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Explain what cells [range, e.g., A1:B10] are calculating. | | |
| 2 | What does the formula in cell [e.g., B15] mean? | | |
| 3 | How is [specific value, e.g., EBITDA] calculated in this range? | | |
| 4 | What are the inputs for the calculation in [cell reference]? | | |
| 5 | Explain the purpose of the table in [range]. | | |
| 6 | What does the conditional formatting in [range] indicate? | | |
| 7 | How does [cell reference] affect other calculations in the model? | | |
| 8 | What is the significance of the value in [cell reference]? | | |
| 9 | Explain the trend shown in [range]. | | |
| 10 | What is the relationship between cells in [range]? | | |
| 11 | How is depreciation calculated in [range or cell]? | | |
| 12 | Explain the amortization schedule in [range]. | | |
| 13 | What does the chart based on [range] represent? | | |
| 14 | How are the percentages in [range] derived? | | |
| 15 | What financial concept is being applied in [range]? | | |
| 16 | Explain the logic behind the IF statement in [cell reference]. | | |
| 17 | How are the lookup functions in [range] being used? | | |
| 18 | What is the purpose of the data validation in [range]? | | |
| 19 | Explain how the pivot table in [range] summarizes the data. | | |
| 20 | What would happen to [cell reference] if [another cell] changed? | | |

### Command Execution Test Cases

The following test cases are designed to validate the add-in's ability to execute various operations in Excel:

#### Basic Cell Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Set cell A1 to 100 | | |
| 2 | Add the formula =SUM(A1:A10) to cell A11 | | |
| 3 | Copy range A1:B10 to C1 | | |
| 4 | Clear the contents of range D1:D10 | | |

#### Formatting Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Format range A1:B10 as currency with red font | | |
| 2 | Add conditional formatting to highlight cells greater than 100 in range C1:C10 | | |
| 3 | Merge cells D1:F1 | | |
| 4 | Format range A1:A10 to be bold and center-aligned | | |

#### Table and Range Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Create a table from data in range A1:D10 with headers | | |
| 2 | Sort range A1:B10 by column A in ascending order | | |
| 3 | Filter range A1:D10 to show only values greater than 50 in column B | | |
| 4 | Add a comment to cell A1 saying "This is the starting value" | | |

#### Worksheet Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Create a new worksheet called "Summary" | | |
| 2 | Set Sheet1 as the active sheet | | |
| 3 | Delete the worksheet named "Temp" | | |
| 4 | Set worksheet zoom to 150% | | |

#### Advanced Features

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Freeze panes at cell B3 | | |
| 2 | Set print area to range A1:H20 on Sheet1 | | |
| 3 | Create a column chart for data in range A1:B10 titled "Sales Report" | | |
| 4 | Format the chart to have a blue background and no gridlines | | |

#### Row and Column Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Set column B width to 15 | | |
| 2 | Group rows 5 to 10 | | |
| 3 | Hide column D | | |
| 4 | Autofit rows 1 to 10 | | |

#### Calculation Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Set calculation mode to manual | | |
| 2 | Enable iterative calculation with 100 iterations | | |
| 3 | Recalculate all worksheets | | |
| 4 | Recalculate range A1:D10 | | |

#### Multiple Operations

| # | Test Query | Result (✅/⚠️/❌) | Feedback |
|---|------------|-----------------|----------|
| 1 | Create a new sheet called Data, add values 1-10 in column A, and create a sum formula in A11 | | |
| 2 | Format range B1:B10 as percentage, add data validation to ensure values are between 0 and 100 | | |
| 3 | Copy data from Sheet1!A1:D10 to Sheet2!A1, format as currency, and create a total row | | |
| 4 | Create a scenario table analyzing cell B10 with input values from A1 | | |

## Testing Notes

1. **Replace placeholders**: When using the test queries, replace placeholders like [Sheet Name], [range], or [cell reference] with actual values from your workbook.

2. **Try variations**: Feel free to modify the suggested queries to better fit your specific financial model.

3. **Response time**: Note if the add-in takes a long time to respond to certain types of questions.

4. **Accuracy**: Pay special attention to whether the explanations are factually correct about the model.

5. **Clarity**: Evaluate how clear and understandable the explanations are, especially for complex concepts.

6. **Command execution**: For command tests, verify that the operations are performed correctly and as expected.

## Submitting Your Results

After completing your testing:

1. Open the HTML file located at C:\Users\shrey\OfficeAddinApps\core-models-excel-addin\docs\testing\testing-form.html in any web browser (just double-click it, or right click and select "Open with" -> "Default Browser")
2. Fill out the test results and feedback for each query
3. Click the "Save Results" button at the bottom
4. Email Shreya the downloaded JSON file (it'll have a .json extension and be populated with some fields inside curly brackets)

Thank you for your help!
