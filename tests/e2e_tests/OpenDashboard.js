/**
 * Simple script to open the test dashboard in a browser window
 */
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');

// Get the absolute path to the dashboard HTML file
const dashboardPath = path.resolve(__dirname, 'TestDashboard.html');

// Ensure the dashboard file exists
if (!fs.existsSync(dashboardPath)) {
  console.error(`Dashboard file not found at: ${dashboardPath}`);
  process.exit(1);
}

// Convert the file path to a URL
const dashboardUrl = `file://${dashboardPath}`;

// Command to open the dashboard in the default browser
const command = process.platform === 'win32' 
  ? `start "" "${dashboardUrl}"` 
  : process.platform === 'darwin' 
    ? `open "${dashboardUrl}"` 
    : `xdg-open "${dashboardUrl}"`;

// Execute the command to open the dashboard
console.log(`Opening dashboard at: ${dashboardUrl}`);
exec(command, (error) => {
  if (error) {
    console.error(`Error opening dashboard: ${error.message}`);
    process.exit(1);
  }
  console.log('Dashboard opened successfully!');
});
