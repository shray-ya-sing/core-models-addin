# Excel Add-in Testing Guide

## Introduction

Thank you for helping test the Excel Financial Models Add-in! This guide will walk you through what the add-in can currently do and provide a structured approach to testing its capabilities.

## Current Capabilities

The add-in is currently in development with the following status:

- ✅ **Workbook Analysis**: The add-in can analyze and explain Excel workbooks, worksheets, and cell ranges
- ✅ **Question Answering**: The add-in can answer questions about financial models and their components
- ✅ **Formula Explanation**: The add-in can explain complex formulas and their purpose
- ❌ **Knowledge Base Integration**: External knowledge integration is not yet functional
- ❌ **Command Execution**: Making modifications to the workbook is still a work in progress

**Important**: At this stage, the add-in is best at "explaining" or "answering questions" about the workbook rather than making modifications. Please focus your testing on these explanation capabilities.

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

The testing is organized into three levels:

1. **Workbook-level**: Questions about the entire financial model
2. **Worksheet-level**: Questions about specific worksheets
3. **Cell Range-level**: Questions about specific cells or ranges

## Test Cases

Some sample questions that might be useful for testing the different scenarios are suggested below:

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

## Testing Notes

1. **Replace placeholders**: When using the test queries, replace placeholders like [Sheet Name], [range], or [cell reference] with actual values from your workbook.

2. **Try variations**: Feel free to modify the suggested queries to better fit your specific financial model.

3. **Response time**: Note if the add-in takes a long time to respond to certain types of questions.

4. **Accuracy**: Pay special attention to whether the explanations are factually correct about the model.

5. **Clarity**: Evaluate how clear and understandable the explanations are, especially for complex concepts.

## Submitting Your Results

After completing your testing:

1. Open the HTML file located at C:\Users\shrey\OfficeAddinApps\core-models-excel-addin\docs\testing\testing-form.html in any web browser (just double-click it, or right click and select "Open with" -> "Default Browser")
2. Fill out the test results and feedback for each query
3. Click the "Save Results" button at the bottom
4. Email Shreya the downloaded JSON file (it'll have a .json extension and be populated with some fields inside curly brackets)

Thank you for your help!
