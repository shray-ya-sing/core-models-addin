# Cori Excel Add-in Installation Guide 

This guide will walk you through installing and running the Excel Financial Models Add-in on your computer. Don't worry if you've never used command-line tools before - we'll go through each step in detail.

## Prerequisites

Before you begin, make sure you have:

1. **Microsoft Excel** installed on your computer (Excel 2016 or newer)
2. **Node.js** - This is required to run the add-in. Download it from [nodejs.org](https://nodejs.org/en/download/) (Download the "LTS" version)
3. **The zip file** containing the add-in that was sent to you

## Step 1: Install Node.js

1. Go to [nodejs.org](https://nodejs.org/en/download/)
2. Download the "LTS" (Long Term Support) version for your operating system:
   - For Windows: Choose the "Windows Installer (.msi)" option
   - For Mac: Choose the "macOS Installer (.pkg)" option

3. Run the installer you downloaded:
   - For Windows: Double-click the .msi file and follow the installation wizard
   - For Mac: Double-click the .pkg file and follow the installation wizard

4. Accept all default settings during installation

5. To verify Node.js was installed correctly:
   - For Windows: Open Command Prompt (search for "cmd" in the Start menu)
   - For Mac: Open Terminal (search for "Terminal" in Spotlight)
   
   Then type this command and press Enter:
   ```
   node --version
   ```
   
   You should see a version number displayed (like v18.16.0)

## Step 2: Extract the Add-in Files

1. Locate the zip file that was sent to you (likely in your Downloads folder)

2. Extract the contents:
   - For Windows: Right-click the zip file and select "Extract All..."
   - For Mac: Double-click the zip file to extract it

3. Choose a location where you want to extract the files:
   - For Windows: A good location is `C:\Projects\core-models-excel-addin`
   - For Mac: A good location is `~/Projects/core-models-excel-addin`

   > Note: If the suggested folders don't exist, you can create them first or choose another location you prefer.

4. After extraction, you should have a folder named `core-models-excel-addin` containing all the add-in files

## Step 3: Install Dependencies

1. Open your command-line tool:
   - For Windows: Open Command Prompt (search for "cmd" in the Start menu)
   - For Mac: Open Terminal (search for "Terminal" in Spotlight)

2. Navigate to the folder where you extracted the add-in:
   - For Windows (if you used the suggested location):
     ```
     cd C:\Projects\core-models-excel-addin
     ```
   - For Mac (if you used the suggested location):
     ```
     cd ~/Projects/core-models-excel-addin
     ```

3. Install the required dependencies by typing this command and pressing Enter:
   ```
   npm install
   ```

4. This process will take several minutes. You'll see lots of text scrolling by - this is normal! Wait until you see a command prompt again, which means the installation is complete.

## Step 4: Start the Add-in Server

1. In the same command-line window, type this command and press Enter:
   ```
   npm run start
   ```

2. You should see messages indicating that the server has started. Keep this window open while using the add-in.

3. If you see any messages about ports being in use, you might need to close other applications or restart your computer before trying again.

## Step 5: Launch the Add-in in Excel

When you run the `npm start` command, the add-in will automatically start and load in Excel:

1. After running `npm start`, Excel will automatically open with a new workbook

2. The add-in should automatically load and appear in the task pane on the right side of Excel

3. If the add-in doesn't load automatically, look for the circular black logo in the top-right corner of Excel that says "Show Task Pane" and click on it

4. If that doesn't work either, look for the button the ribbon that says "Add-ins", then go to "Developer Add-ins" and a template black circular logo with the subtext core-models-excel-addin should be there. Click on it to load the addin.

## Using the Add-in

Once the add-in is loaded, you'll see the Financial Models interface in the task pane. You can now:

1. Use natural language to ask questions about your Excel workbook
2. Request modifications to your financial models
3. Get assistance with formulas and data analysis

## Troubleshooting

If you encounter any issues:

1. **Add-in doesn't appear in Excel**:
   - Make sure the command prompt window running the server is still open
   - Try closing Excel completely and reopening it
   - Verify that you selected the correct `manifest.xml` file

2. **Server won't start**:
   - Make sure you're in the correct folder when running the commands
   - Try restarting your computer and trying again
   - Check that Node.js was installed correctly

3. **Add-in loads but doesn't work properly**:
   - Make sure you completed all the installation steps
   - Check that the command prompt window is still running the server
   - Try refreshing the add-in by closing and reopening the task pane

## Closing the Add-in

When you're done using the add-in:

1. You can close Excel normally

2. To stop the server, go back to the command prompt window and press `Ctrl+C` (on both Windows and Mac)

3. When asked if you want to terminate the batch job, type `Y` and press Enter

The next time you want to use the add-in, you'll need to start the server again (Step 4), but you won't need to reinstall everything.
