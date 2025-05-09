@echo off
echo Starting Test Results Server...
start "Test Results Server" node "%~dp0TestResultsServer.js"
timeout /t 2 /nobreak > nul
echo Opening Excel Test Dashboard...
start "" "http://localhost:3001"
echo Dashboard opened in your default browser.
