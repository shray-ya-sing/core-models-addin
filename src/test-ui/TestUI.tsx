/**
 * Test UI Component
 * A minimalist UI for running integration tests
 */
import * as React from "react";
import { useState, useEffect } from "react";
import { ClientQueryProcessor } from "../client/services/request-processing/ClientQueryProcessor";
import { ClientExcelCommandInterpreter } from "../client/services/actions/ClientExcelCommandInterpreter";
import { TestRunner } from "../../tests/e2e_tests/TestRunner";
import { AllTestScenarios, TestResult, TestScenario } from "../../tests/e2e_tests/AllTestScenarios";

// Styles for the test UI
const styles = {
  container: {
    fontFamily: "Arial, 'Helvetica Neue', Helvetica, sans-serif",
    fontSize: "clamp(0.75rem, 1vw, 0.875rem)",
    padding: "0.5rem",
    backgroundColor: "#1a1a1a",
    color: "#e0e0e0",
    borderTop: "1px solid #333",
    maxHeight: "30vh",
    overflow: "auto"
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "0.5rem"
  },
  title: {
    margin: 0,
    fontWeight: "normal",
    fontSize: "clamp(0.875rem, 1.2vw, 1rem)",
    color: "#aaa"
  },
  button: {
    backgroundColor: "#2d3748",
    color: "#e2e8f0",
    border: "none",
    padding: "0.25rem 0.5rem",
    borderRadius: "0.25rem",
    fontSize: "clamp(0.75rem, 1vw, 0.875rem)",
    cursor: "pointer",
    fontFamily: "inherit"
  },
  results: {
    marginTop: "0.5rem",
    fontSize: "clamp(0.7rem, 0.9vw, 0.8rem)"
  },
  success: {
    color: "#68d391"
  },
  failure: {
    color: "#fc8181"
  },
  step: {
    marginLeft: "0.5rem",
    marginBottom: "0.25rem"
  }
};

interface TestUIProps {
  queryProcessor: ClientQueryProcessor;
  commandInterpreter: ClientExcelCommandInterpreter;
  onClose?: () => void;
}

/**
 * TestUI Component
 * A minimalist UI for running integration tests
 */
export const TestUI: React.FC<TestUIProps> = ({
  queryProcessor,
  commandInterpreter,
  onClose
}) => {
  const [testResult, setTestResult] = useState<TestResult | null>(null);
  const [isRunning, setIsRunning] = useState(false);
  const [selectedScenario, setSelectedScenario] = useState<TestScenario | null>(null);
  const [testResults, setTestResults] = useState<TestResult[]>([]);

  // Initialize test runner when component mounts
  useEffect(() => {
    console.log('TestUI component initialized');
  }, []);

  // Function to open the dashboard
  const openDashboard = () => {
    try {
      // Get the current directory
      const currentDir = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/'));
      
      // Construct the path to the dashboard HTML file
      const dashboardPath = `${currentDir}/tests/e2e_tests/TestDashboard.html`;
      
      // Open the dashboard in a new window
      const dashboardWindow = window.open(dashboardPath, '_blank', 'width=1000,height=800');
      
      if (!dashboardWindow) {
        console.error('Failed to open dashboard window. Please check if pop-ups are blocked.');
      }
    } catch (error) {
      console.error('Error opening dashboard:', error);
    }
  };

  // Function to run a specific test scenario
  const runTestScenario = async (scenario: TestScenario) => {
    if (isRunning) return;
    
    setIsRunning(true);
    setSelectedScenario(scenario);
    setTestResult(null);
    
    try {
      // Create a new test runner
      const testRunner = new TestRunner(queryProcessor, commandInterpreter);
      
      // Open the dashboard
      openDashboard();
      
      // Run the test
      console.log(`Running test scenario: ${scenario.name}`);
      const result = await testRunner.runTest(scenario);
      
      // Store the results
      const updatedResults = [...testResults, result];
      setTestResults(updatedResults);
      
      // Store the results in localStorage
      localStorage.setItem('excelTestResults', JSON.stringify(updatedResults));
      
      // Update the state
      setIsRunning(false);
      setTestResult(result);
    } catch (error) {
      console.error(`Error running test scenario ${scenario.name}:`, error);
      setIsRunning(false);
    }
  };
  
  // Function to run all test scenarios
  const runAllTests = async () => {
    if (isRunning) return;
    
    setIsRunning(true);
    setTestResults([]);
    
    try {
      // Create a new test runner
      const testRunner = new TestRunner(queryProcessor, commandInterpreter);
      
      // Open the dashboard
      openDashboard();
      
      // Run all tests
      const results: TestResult[] = [];
      
      for (const scenario of AllTestScenarios) {
        console.log(`Running test scenario: ${scenario.name}`);
        setSelectedScenario(scenario);
        
        const result = await testRunner.runTest(scenario);
        results.push(result);
        
        // Update results after each test
        setTestResults([...results]);
      }
      
      // Store the results in localStorage
      localStorage.setItem('excelTestResults', JSON.stringify(results));
      
      // Update the state
      setIsRunning(false);
      setSelectedScenario(null);
    } catch (error) {
      console.error('Error running all tests:', error);
      setIsRunning(false);
      setSelectedScenario(null);
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.header}>
        <h3 style={styles.title}>Test Mode</h3>
        <div>
          <button 
            style={styles.button} 
            onClick={runAllTests}
            disabled={isRunning}
          >
            {isRunning ? "Running..." : "Run All Tests"}
          </button>
          <button 
            style={{...styles.button, marginLeft: "0.5rem"}} 
            onClick={openDashboard}
          >
            Dashboard
          </button>
          {onClose && (
            <button 
              style={{...styles.button, marginLeft: "0.5rem"}} 
              onClick={onClose}
            >
              Close
            </button>
          )}
        </div>
      </div>
      
      <div style={{marginTop: '0.5rem', marginBottom: '0.5rem'}}>
        <div>Available Test Scenarios:</div>
        <div style={{display: 'flex', flexWrap: 'wrap', gap: '0.5rem', marginTop: '0.5rem'}}>
          {AllTestScenarios.map(scenario => (
            <button 
              key={scenario.id} 
              style={{
                ...styles.button, 
                fontSize: 'clamp(0.65rem, 0.8vw, 0.75rem)',
                backgroundColor: selectedScenario?.id === scenario.id ? '#4a5568' : '#2d3748'
              }}
              onClick={() => runTestScenario(scenario)}
              disabled={isRunning}
            >
              {scenario.name}
            </button>
          ))}
        </div>
      </div>
      
      {isRunning && selectedScenario && (
        <div>Running test scenario: {selectedScenario.name}...</div>
      )}
      
      {testResults.length > 0 && (
        <div style={styles.results}>
          <div style={{marginBottom: '0.5rem', fontWeight: 'bold'}}>
            Test Results: {testResults.filter(r => r.success).length}/{testResults.length} Passed
          </div>
          
          {testResults.map(result => (
            <div key={result.scenarioId} style={{marginBottom: '1rem'}}>
              <div style={result.success ? styles.success : styles.failure}>
                {result.name}: {result.success ? "PASSED" : "FAILED"} 
                ({result.duration}ms)
              </div>
              
              {/* Show steps for the most recent test result */}
              {testResult && testResult.scenarioId === result.scenarioId && (
                <div style={{marginLeft: '1rem'}}>
                  {result.steps.map(step => (
                    <div key={step.stepId} style={styles.step}>
                      <span style={step.success ? styles.success : styles.failure}>
                        {step.success ? "✓" : "✗"}
                      </span>
                      {" "}
                      {step.query}
                      {step.message && !step.success && (
                        <div style={{marginLeft: "1rem", color: "#fc8181"}}>
                          {step.message}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default TestUI;
