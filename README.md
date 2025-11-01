# CoreModels Addin

The CoreModels Addin is a fully client-side Excel assistant that helps analysts move faster through their Excel workflow by extracting data, explaining, question answering, and executing tasks—all directly in the browser.

## Technology Stack

### Client-side Architecture
- **React**: UI framework for building the add-in interface
- **TypeScript**: Type-safe JavaScript for improved development experience
- **Office.js**: Excel JavaScript API for interacting with Excel
- **Fluent UI**: Microsoft's design system for Office add-ins
- **Anthropic Claude Integration**: AI service for natural language processing and financial model understanding
- **WebSockets**: For real-time command execution updates
- **Client-side State Management**: In-browser workbook state cache with granular sheet-level tracking

### Storage
- **LocalStorage**: Browser-based storage for user preferences and state
- **Browser Cache**: Client-side storage for workbook snapshots and query results

### Development & Testing
- **Jest**: Testing framework for unit and integration tests
- **Webpack**: Module bundler for building the application
- **ESLint**: Code linting for maintaining code quality

## How to run this project

### Understanding Excel Add-ins

Excel add-ins are web applications that extend Excel's functionality using the Office JavaScript API. They consist of two main parts:

1. **Web Application**: A web app built with HTML, CSS, and JavaScript/TypeScript that runs in a browser control or iframe within Excel. This provides the UI and business logic.

2. **Manifest File**: An XML file that specifies how the add-in should be integrated into Excel, including permissions, UI elements, and entry points.

This CoreModels Addin uses a 100% client-side architecture:

When the add-in runs, it:
1. **Loads the web application** in a task pane within Excel
2. **Initializes client-side services** for workbook state management, command execution, and LLM integration
3. **Interacts with Excel** using the Office.js API to read and write data
4. **Processes commands** entirely in the browser using the client-side implementation
5. **Uses intelligent caching** to minimize redundant workbook state captures
6. **Performs granular workbook analysis** with sheet-level dependency tracking
7. **Connects to external APIs** for LLM processing only when needed

### Prerequisites

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify that you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Office connected to a Microsoft 365 subscription (for production use)
- Anthropic API key (for LLM integration)

### Run the add-in using Office Add-ins Development Kit extension

1. **Open the Office Add-ins Development Kit**
    
    In the **Activity Bar**, select the **Office Add-ins Development Kit** icon to open the extension.

1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the Quick Pick menu, select the option **Excel Desktop (Edge Chromium)**.

    The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Excel desktop application launches and sideloads the add-in.

1. **Stop Previewing Your Office Add-in**

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

### Environment Variables
Create a `.env` file in the root directory with the following variables:

```
VOYAGEAI_API_KEY=your_api_key_here
USE_MULTIMODAL_EMBEDDINGS=true
ANTHROPIC_API_KEY=your_api_key_here
```

## Troubleshooting

If you have problems running the add-in, take these steps.

- Close any open instances of Excel.
- Close the previous web server started for the add-in with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.
- Check that your VoyageAI API key is correctly set in the `.env` file.
- Check that your Anthropic API key is correctly set in the .env file.
- Verify that the necessary npm packages are installed with `npm install`.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the add-in on Excel on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).


## Copyright and License

© Shreya Singh. All rights reserved.

This software and its documentation are proprietary to Shreya Singh. The software is provided under a license agreement containing restrictions on use, duplication, disclosure, and is protected by intellectual property laws. This software may not be used, copied, distributed, modified, or disclosed to any third party without the express written permission of Shreya Singh.
