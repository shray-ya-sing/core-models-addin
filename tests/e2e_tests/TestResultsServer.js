/**
 * Simple server to serve the test dashboard and receive test results
 */
const express = require('express');
const path = require('path');
const cors = require('cors');
const fs = require('fs');

const app = express();
const PORT = 3001;

// Enable CORS for all routes
app.use(cors());

// Parse JSON bodies
app.use(express.json());

// Serve static files from the e2e_tests directory
app.use(express.static(path.join(__dirname)));

// Store the latest test results
let latestResults = null;

// Endpoint to receive test results
app.post('/api/test-results', (req, res) => {
  console.log('Received test results:', JSON.stringify(req.body, null, 2));
  latestResults = req.body;
  
  // Also save to a file for persistence
  fs.writeFileSync(
    path.join(__dirname, 'latest-results.json'), 
    JSON.stringify(latestResults, null, 2)
  );
  
  res.json({ success: true });
});

// Endpoint to get the latest test results
app.get('/api/test-results', (req, res) => {
  if (latestResults) {
    res.json(latestResults);
  } else {
    // Try to read from file if available
    try {
      const filePath = path.join(__dirname, 'latest-results.json');
      if (fs.existsSync(filePath)) {
        const fileData = fs.readFileSync(filePath, 'utf8');
        latestResults = JSON.parse(fileData);
        res.json(latestResults);
      } else {
        res.json({ message: 'No test results available yet' });
      }
    } catch (error) {
      console.error('Error reading test results file:', error);
      res.json({ message: 'No test results available yet' });
    }
  }
});

// Endpoint to save test results to a file in the results directory
app.post('/api/save-test-results', (req, res) => {
  try {
    const { filename, content } = req.body;
    
    if (!filename || !content) {
      return res.status(400).json({ error: 'Filename and content are required' });
    }
    
    // Create the results directory if it doesn't exist
    const resultsDir = path.join(__dirname, 'results');
    if (!fs.existsSync(resultsDir)) {
      fs.mkdirSync(resultsDir, { recursive: true });
      console.log(`Created results directory: ${resultsDir}`);
    }
    
    // Save the file
    const filepath = path.join(resultsDir, filename);
    fs.writeFileSync(filepath, content);
    
    console.log(`Saved test results to: ${filepath}`);
    
    // Also update the latest results
    latestResults = JSON.parse(content);
    
    // Return success
    res.json({ success: true, filepath });
  } catch (error) {
    console.error('Error saving test results:', error);
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to save test results
app.post('/api/save-results', (req, res) => {
  try {
    const testResults = req.body;
    
    if (!testResults) {
      return res.status(400).json({ error: 'No test results provided' });
    }
    
    // Get the results directory path
    const resultsDir = path.join(__dirname, 'results');
    
    // Create the directory if it doesn't exist
    if (!fs.existsSync(resultsDir)) {
      fs.mkdirSync(resultsDir, { recursive: true });
      console.log(`Created results directory: ${resultsDir}`);
    }
    
    // Generate a filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    
    // Check if this is an aggregated result or a single test result
    let filename;
    if (testResults.results && testResults.summary) {
      // This is an aggregated result
      filename = `aggregated-results-${timestamp}.json`;
    } else {
      // This is a single test result
      filename = `test-results-${timestamp}.json`;
    }
    
    const filepath = path.join(resultsDir, filename);
    
    // Write the results to a file
    fs.writeFileSync(filepath, JSON.stringify(testResults, null, 2));
    
    console.log(`Test results saved to: ${filepath}`);
    res.json({ success: true, filename });
  } catch (error) {
    console.error('Error saving test results:', error);
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to get a list of result files
app.get('/api/result-files', (req, res) => {
  try {
    // Get the results directory path
    const resultsDir = path.join(__dirname, 'results');
    
    // Create the directory if it doesn't exist
    if (!fs.existsSync(resultsDir)) {
      fs.mkdirSync(resultsDir, { recursive: true });
      console.log(`Created results directory: ${resultsDir}`);
      return res.json({ files: [] });
    }
    
    // Get all JSON files in the directory
    const files = fs.readdirSync(resultsDir)
      .filter(file => file.endsWith('.json'))
      .sort()
      .reverse(); // Most recent first
    
    res.json({ files });
  } catch (error) {
    console.error('Error getting result files:', error);
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to get a specific result file
app.get('/api/result-file/:filename', (req, res) => {
  try {
    const { filename } = req.params;
    
    if (!filename) {
      return res.status(400).json({ error: 'Filename is required' });
    }
    
    // Get the file path
    const filepath = path.join(__dirname, 'results', filename);
    
    // Check if the file exists
    if (!fs.existsSync(filepath)) {
      return res.status(404).json({ error: 'File not found' });
    }
    
    // Read the file
    const fileContent = fs.readFileSync(filepath, 'utf8');
    const resultData = JSON.parse(fileContent);
    
    // Return the file content
    res.json(resultData);
  } catch (error) {
    console.error('Error getting result file:', error);
    res.status(500).json({ error: error.message });
  }
});

// Default route serves the dashboard
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'TestDashboard.html'));
});

// Start the server
app.listen(PORT, () => {
  console.log(`Test dashboard server running at http://localhost:${PORT}`);
  console.log(`Open http://localhost:${PORT} in your browser to view the dashboard`);
});
