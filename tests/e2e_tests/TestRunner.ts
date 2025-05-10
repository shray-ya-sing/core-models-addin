/**
 * Test Runner
 * Executes test scenarios and captures results
 */
import { ClientQueryProcessor } from "../../src/client/services/request-processing/ClientQueryProcessor";
import { ClientExcelCommandInterpreter } from "../../src/client/services/actions/ClientExcelCommandInterpreter";
import { 
  AllTestScenarios, 
  ExpectedOutcome, 
  SampleDataScenario, 
  TestResult, 
  TestScenario, 
  TestStep, 
  TestStepResult 
} from "./AllTestScenarios";
import { 
  ExpectedCellFormat, 
  ExpectedChartFormat, 
  FormatVerifier, 
  FormatVerificationResult 
} from "./FormatVerifier";
// Use the global Office and Excel objects instead of importing them
// This avoids TypeScript errors with the office-js module
declare const Excel: any;
declare const Office: any;

/**
 * Test runner for executing integration tests
 */
export class TestRunner {
  private queryProcessor: ClientQueryProcessor;
  private commandInterpreter: ClientExcelCommandInterpreter;
  private scenario: TestScenario | null = null;
  private currentTestResult: TestResult | null = null;
  private testListeners: Array<(result: TestResult) => void> = [];

  /**
   * Constructor
   * @param queryProcessor The query processor to test
   * @param commandInterpreter The command interpreter to test
   */
  constructor(
    queryProcessor: ClientQueryProcessor,
    commandInterpreter: ClientExcelCommandInterpreter
  ) {
    this.queryProcessor = queryProcessor;
    this.commandInterpreter = commandInterpreter;
  }

  /**
   * Add a test result listener
   * @param listener The listener function
   */
  addTestListener(listener: (result: TestResult) => void): void {
    this.testListeners.push(listener);
  }

  /**
   * Run a test scenario
   * @param scenario The test scenario to run
   * @returns The test result
   */
  public async runTest(scenario: TestScenario = SampleDataScenario): Promise<TestResult> {
    console.log(`Running test scenario: ${scenario.name}`);
    
    // Clear any existing charts and data from previous tests
    await this.clearWorkbook();
    
    // Setup any required worksheets for the test scenario
    await this.setupRequiredWorksheets(scenario);
    
    this.scenario = scenario;
    
    // Initialize test result
    const startTime = Date.now();
    this.currentTestResult = {
      scenarioId: scenario.id,
      name: scenario.name,
      success: true,
      steps: [],
      startTime,
      endTime: 0,
      duration: 0
    };
    
    try {
      // Run each step in the scenario
      for (const step of scenario.steps) {
        const stepResult = await this.runTestStep(step);
        
        // Add step result to test result
        this.currentTestResult.steps.push(stepResult);
        
        // If step failed, mark test as failed
        if (!stepResult.success) {
          this.currentTestResult.success = false;
        }
      }
      
      // Clean up after the test is complete
      await this.clearWorkbook();
      
      // Complete the test result
      const endTime = Date.now();
      this.currentTestResult.endTime = endTime;
      this.currentTestResult.duration = endTime - startTime;
      
      // Notify listeners of test results
      this.notifyListeners(this.currentTestResult);
      
      return this.currentTestResult;
    } catch (error) {
      console.error(`Error running test scenario: ${error.message}`);
      this.currentTestResult.success = false;
      
      // Complete the test result
      const endTime = Date.now();
      this.currentTestResult.endTime = endTime;
      this.currentTestResult.duration = endTime - startTime;
      
      // Notify listeners of test results
      this.notifyListeners(this.currentTestResult);
      
      return this.currentTestResult;
    }
  }

  // clearWorkbook is implemented below

  /**
   * Run a single test step
   * @param step The test step to run
   * @returns The test step result
   */
  private async runTestStep(step: TestStep): Promise<TestStepResult> {
    console.log(`Running test step: ${step.id} - ${step.description}`);
    
    try {
      // Process the query
      const result = await this.queryProcessor.processQuery(step.query);
      // Add a delay here to ensure Excel operations have fully completed
      console.log('Waiting for Excel operations to complete...');
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Verify the expected outcome
      const verificationResult = await this.verifyOutcome(step.expectedOutcome);
      
      return {
        stepId: step.id,
        query: step.query,
        success: verificationResult.success,
        message: verificationResult.message,
        details: verificationResult.details
      };
    } catch (error) {
      console.error(`Error running test step: ${error.message}`);
      
      return {
        stepId: step.id,
        query: step.query,
        success: false,
        message: `Error: ${error.message}`
      };
    }
  }

  /**
   * Verify the expected outcome
   * @param expectedOutcome The expected outcome
   * @returns The verification result
   */
  private async verifyOutcome(expectedOutcome: ExpectedOutcome): Promise<{
    success: boolean;
    message: string;
    details: any;
  }> {
    const details: any = {};
    let success = true;
    let message = "All verifications passed";
    
    try {
      // Verify cell values if specified
      if (expectedOutcome.cellValues && expectedOutcome.cellValues.length > 0) {
        const cellValueMatches = await this.verifyCellValues(expectedOutcome.cellValues);
        details.cellValueMatches = cellValueMatches;
        
        // Check if all cell values match
        const allCellValuesMatch = cellValueMatches.every(match => match.match);
        if (!allCellValuesMatch) {
          success = false;
          message = "Cell value verification failed";
        }
      }
      
      // Verify cell formatting if specified
      if (expectedOutcome.cellFormatting && expectedOutcome.cellFormatting.length > 0) {
        // Convert the expected formatting to the format required by FormatVerifier
        const cellFormats: ExpectedCellFormat[] = expectedOutcome.cellFormatting.map(format => ({
          sheetName: format.sheetName,
          cellAddress: format.cellAddress,
          properties: format.properties
        }));
        
        const formatMatches = await FormatVerifier.verifyCellFormatting(cellFormats);
        details.cellFormatMatches = formatMatches;
        
        // Check if all formatting properties match
        const allFormatMatches = formatMatches.every(match => match.match);
        if (!allFormatMatches) {
          success = false;
          message = success ? "Cell formatting verification failed" : message;
        }
      }
      
      // Verify chart formatting if specified
      if (expectedOutcome.chartFormatting && expectedOutcome.chartFormatting.length > 0) {
        // Convert the expected formatting to the format required by FormatVerifier
        const chartFormats: ExpectedChartFormat[] = expectedOutcome.chartFormatting.map(format => ({
          sheetName: format.sheetName,
          chartName: format.chartName,
          chartIndex: format.chartIndex,
          properties: format.properties
        }));
        
        const chartFormatMatches = await FormatVerifier.verifyChartFormatting(chartFormats);
        details.chartFormatMatches = chartFormatMatches;
        
        // Check if all chart formatting properties match
        const allChartFormatMatches = chartFormatMatches.every(match => match.match);
        if (!allChartFormatMatches) {
          success = false;
          message = success ? "Chart formatting verification failed" : message;
        }
      }
      
      // Verify operations if specified
      if (expectedOutcome.expectedOperations && expectedOutcome.expectedOperations.length > 0) {
        // For now, we just log the expected operations
        // In the future, this could be enhanced to verify the operations were actually executed
        console.log("Expected operations:", expectedOutcome.expectedOperations);
      }
      
      return { success, message, details };
    } catch (error) {
      return {
        success: false,
        message: `Verification error: ${error.message}`,
        details
      };
    }
  }

  /**
   * Verify cell values
   * @param expectedCellValues The expected cell values
   * @returns The verification results
   */
  private async verifyCellValues(expectedCellValues: Array<{
    sheetName: string;
    cellAddress: string;
    expectedValue: any;
  }>): Promise<Array<{
    address: string;
    expected: any;
    actual: any;
    match: boolean;
  }>> {
    const results: Array<{
      address: string;
      expected: any;
      actual: any;
      match: boolean;
    }> = [];
    
    console.log('Verifying cell values:', expectedCellValues);
    
    await Excel.run(async (context) => {
      for (const cellValue of expectedCellValues) {
        try {
          const sheet = context.workbook.worksheets.getItem(cellValue.sheetName);
          const range = sheet.getRange(cellValue.cellAddress);
          range.load("values");
          
          await context.sync();
          
          const actualValue = range.values[0][0];
          
          // Determine if we need to compare as numbers or strings
          let match = false;
          
          if (typeof cellValue.expectedValue === 'number') {
            // For numeric values, convert actual value to number for comparison
            // This handles formatted numbers like "$120,000.00" -> 120000
            const actualNumber = typeof actualValue === 'number' 
              ? actualValue 
              : Number(String(actualValue).replace(/[^\d.-]/g, ''));
            
            // Compare as numbers
            match = !isNaN(actualNumber) && actualNumber === cellValue.expectedValue;
            console.log(`Numeric comparison for ${cellValue.sheetName}!${cellValue.cellAddress} - Expected: ${cellValue.expectedValue}, Actual: ${actualValue} (${actualNumber}), Match: ${match}`);
          } else {
            // For strings, do case-insensitive comparison
            const expectedStr = String(cellValue.expectedValue).trim().toLowerCase();
            const actualStr = String(actualValue).trim().toLowerCase();
            match = expectedStr === actualStr;
            console.log(`String comparison for ${cellValue.sheetName}!${cellValue.cellAddress} - Expected: "${expectedStr}", Actual: "${actualStr}", Match: ${match}`);
          }
          
          // Log statement is handled in the conditional blocks above
          
          results.push({
            address: `${cellValue.sheetName}!${cellValue.cellAddress}`,
            expected: cellValue.expectedValue,
            actual: actualValue,
            match
          });
        } catch (error) {
          console.error(`Error verifying cell ${cellValue.sheetName}!${cellValue.cellAddress}:`, error);
          results.push({
            address: `${cellValue.sheetName}!${cellValue.cellAddress}`,
            expected: cellValue.expectedValue,
            actual: null,
            match: false
          });
        }
      }
    });
    
    return results;
  }

  /**
   * Setup required worksheets for a test scenario
   * This ensures that worksheets needed for the test exist before running the test
   * @param scenario The test scenario to setup worksheets for
   */
  private async setupRequiredWorksheets(scenario: TestScenario): Promise<void> {
    console.log('Setting up required worksheets for test scenario');
    
    // Check if this is the worksheet management scenario
    if (scenario.id === 'worksheet-mgmt-001') {
      try {
        await Excel.run(async (context) => {
          // For the worksheet management scenario, we need to ensure Sheet1 exists
          const worksheets = context.workbook.worksheets;
          worksheets.load("items/name");
          await context.sync();
          
          // Check if Sheet1 exists
          let sheet1Exists = false;
          for (const ws of worksheets.items) {
            if (ws.name === "Sheet1") {
              sheet1Exists = true;
              break;
            }
          }
          
          // If Sheet1 doesn't exist, add it
          if (!sheet1Exists) {
            worksheets.add("Sheet1");
            console.log("Added Sheet1 for worksheet management test");
          }
          
          await context.sync();
        });
      } catch (error) {
        console.error('Error setting up worksheets:', error);
      }
    }
    
    // For financial modeling scenario, ensure required worksheets exist
    if (scenario.id === 'financial-modeling-001') {
      try {
        await Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          worksheets.load("items/name");
          await context.sync();
          
          // Check for required worksheets and add if missing
          const requiredSheets = ["Income Statement", "Balance Sheet"];
          const existingSheetNames = worksheets.items.map(ws => ws.name);
          
          for (const sheetName of requiredSheets) {
            if (!existingSheetNames.includes(sheetName)) {
              worksheets.add(sheetName);
              console.log(`Added required worksheet: ${sheetName}`);
            }
          }
          
          await context.sync();
        });
      } catch (error) {
        console.error('Error setting up financial worksheets:', error);
      }
    }
  }
  
  /**
   * Clear all charts and data from the workbook
   * This ensures each test starts with a clean slate
   */
  private async clearWorkbook(): Promise<void> {
    console.log('Clearing workbook for clean test environment');
    
    try {
      await Excel.run(async (context) => {
        // Get all worksheets
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        
        // Process each worksheet
        for (const worksheet of worksheets.items) {
          console.log(`Clearing worksheet: ${worksheet.name}`);
          
          // Clear all charts
          const charts = worksheet.charts;
          charts.load("items/name");
          await context.sync();
          
          console.log(`Found ${charts.items.length} charts to remove in ${worksheet.name}`);
          
          // Delete each chart
          for (const chart of charts.items) {
            console.log(`Removing chart: ${chart.name || 'unnamed chart'}`);
            chart.delete();
          }
          
          // Clear all data
          const usedRange = worksheet.getUsedRange();
          usedRange.load("address");
          await context.sync();
          
          // Only clear if there is data
          if (usedRange && usedRange.address) {
            console.log(`Clearing data in range: ${usedRange.address}`);
            usedRange.clear("All");
          }
        }
        
        await context.sync();
        console.log('Workbook cleared successfully');
      });
    } catch (error) {
      console.error('Error clearing workbook:', error);
    }
  }

  /**
   * Notify listeners of test results
   * @param result The test result
   */
  private notifyListeners(result: TestResult): void {
    // Send results to any registered listeners
    for (const listener of this.testListeners) {
      try {
        listener(result);
      } catch (error) {
        console.error(`Error notifying test listener: ${error.message}`);
      }
    }
    
    // Send results to the dashboard
    this.sendResultsToDashboard(result);
  }
  
  /**
   * Send test results to the dashboard
   * @param result The test result
   */
  private async sendResultsToDashboard(testResults: TestResult) {
    try {
      // Add the original test query and description to the results
      const testQuery = this.scenario?.steps[0]?.query || 'Unknown test query';
      const enrichedResults = {
        ...testResults,
        testQuery: testQuery,
        testDescription: this.scenario?.description || 'Unknown test description'
      };
      
      // Generate a filename with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const filename = `test_results_${timestamp}.json`;
      const filepath = `./tests/e2e_tests/results/${filename}`;
      
      // Convert the test results to a JSON string
      const testResultsJson = JSON.stringify(enrichedResults, null, 2);
      
      // Save to localStorage as a backup
      localStorage.setItem('excelTestResults', testResultsJson);
      localStorage.setItem('lastTestResultFile', filename);
      
      // Save to file using Office.js file download
      // Create a blob with the JSON data
      const blob = new Blob([testResultsJson], { type: 'application/json' });
      
      // Create a URL for the blob
      const url = URL.createObjectURL(blob);
      
      // Create a link element
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      
      // Append the link to the document
      document.body.appendChild(a);
      
      // Click the link to download the file
      a.click();
      
      // Clean up
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      console.log(`Test results saved to file: ${filename}`);
      console.log(`Please manually move the downloaded file to: ${filepath}`);
      
    } catch (error) {
      console.error('Error saving test results to file:', error);
      // Fallback to localStorage only
      localStorage.setItem('excelTestResults', JSON.stringify(testResults));
    }
  }
}
