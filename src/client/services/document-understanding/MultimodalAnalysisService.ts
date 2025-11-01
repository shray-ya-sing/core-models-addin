/**
 * Service for multimodal analysis of Excel workbooks
 * Integrates with LLMs for formatting analysis
 */
import { performMultimodalAnalysis } from './WorkbookUtils';
import { FormattingProtocolAnalyzer } from './FormattingProtocolAnalyzer';
import { FormattingProtocol, WorkbookFormattingMetadata } from './FormattingModels';
import { ClientAnthropicService } from '../llm/ClientAnthropicService';

// Types for the multimodal analysis options
export interface MultimodalAnalysisOptions {
  sheets?: string[];
  charts?: Array<{ sheetName: string; chartIndex: number }>;
  ranges?: Array<{ sheetName: string; range: string }>;
}

// Configuration
const DEFAULT_IMAGE_CONVERSION_ENDPOINT = 'http://localhost:8080/api/ExcelImage/export';

// Types
export interface MultimodalAnalysisResult {
  images: string[];
  formattingAnalysis?: FormattingProtocol;
  formattingMetadata?: WorkbookFormattingMetadata;
  metadata: {
    timestamp: string;
    sheetCount: number;
    options?: MultimodalAnalysisOptions;
  };
}

/**
 * Service for performing multimodal analysis on Excel workbooks
 */
export class MultimodalAnalysisService {
  private apiEndpoint: string;
  private analysisResults: Map<string, MultimodalAnalysisResult> = new Map();
  private formattingProtocolAnalyzer: FormattingProtocolAnalyzer | null = null;
  
  // Track formatting protocols per workbook
  private workbookFormattingProtocols: Map<string, {
    formattingProtocol: FormattingProtocol;
    formattingMetadata: WorkbookFormattingMetadata;
    timestamp: string;
  }> = new Map();
  
  // Current workbook tracking
  private currentWorkbookId: string = '';
  
  // Legacy cache variables - kept for backward compatibility
  private cachedFormattingProtocol: FormattingProtocol | null = null;
  private cachedFormattingMetadata: WorkbookFormattingMetadata | null = null;
  private lastAnalysisTimestamp: string = '';

  /**
   * Creates a new instance of the MultimodalAnalysisService
   * @param apiEndpoint Optional custom API endpoint for image conversion
   * @param anthropicService Optional Anthropic service for LLM analysis
   */
  constructor(
    apiEndpoint: string = DEFAULT_IMAGE_CONVERSION_ENDPOINT,
    anthropicService?: ClientAnthropicService
  ) {
    this.apiEndpoint = apiEndpoint;
    
    // Initialize the formatting protocol analyzer if Anthropic service is provided
    if (anthropicService) {
      this.formattingProtocolAnalyzer = new FormattingProtocolAnalyzer(anthropicService);
    }

  }
  
  /**
   * Sets the current workbook ID and ensures its formatting is analyzed
   * @param workbookId The ID of the current workbook
   * @param waitForCompletion Whether to wait for the full analysis to complete (default: false)
   * @returns Promise that resolves when the workbook's formatting has been analyzed
   */
  public async setWorkbookAndEnsureFormatting(workbookId: string, waitForCompletion: boolean = false): Promise<void> {
    // If this is a different workbook than the current one
    if (this.currentWorkbookId !== workbookId) {
      console.log(`
=======================================================
🔄 SETTING WORKBOOK AND ENSURING FORMATTING: ${workbookId}
=======================================================`);
      console.log(`Wait for completion: ${waitForCompletion ? 'YES' : 'NO'}`);
      this.currentWorkbookId = workbookId;
      const hasExistingFormatting = this.hasWorkbookFormatting(workbookId);
      
      if (!hasExistingFormatting) {
        console.log(`🔍 Starting formatting protocol analysis for workbook: ${workbookId}`);
        
        try {
          console.log(`Step 1: Calling analyzeFormattingProtocol with waitForCompletion=${waitForCompletion}`);
          // Analyze the formatting protocol for this workbook
          // If waitForCompletion is true, this will wait for the full analysis including LLM processing
          const formattingProtocol = await this.analyzeFormattingProtocol(waitForCompletion);
          
          console.log(`Step 2: Received formatting protocol with ${Object.keys(formattingProtocol).length} top-level categories`);
          console.log(`   Protocol contains: ${Object.keys(formattingProtocol).join(', ')}`);
          
          // Store the formatting protocol for this workbook
          console.log(`Step 3: Storing formatting protocol in cache for workbook: ${workbookId}`);
          const formattingData = {
            formattingProtocol,
            formattingMetadata: this.cachedFormattingMetadata!,
            timestamp: new Date().toISOString()
          };
          
          this.workbookFormattingProtocols.set(workbookId, formattingData);
          
          console.log(`✅ Formatting protocol analysis ${waitForCompletion ? 'fully' : 'initially'} complete for workbook: ${workbookId}`);
        } catch (error) {
          console.error(`❌ ERROR ANALYZING FORMATTING PROTOCOL FOR WORKBOOK ${workbookId}:`, error);
          console.error(`Error stack trace: ${error.stack}`);
          throw error;
        }
      } else {
        console.log(`📋 Using cached formatting protocol for workbook: ${workbookId}`);
        
        // Update the legacy cache variables for backward compatibility
        const workbookData = this.workbookFormattingProtocols.get(workbookId)!;
        this.cachedFormattingProtocol = workbookData.formattingProtocol;
        this.cachedFormattingMetadata = workbookData.formattingMetadata;
        this.lastAnalysisTimestamp = workbookData.timestamp;
      }
    }
  }
  
  /**
   * Gets the formatting protocol for the current workbook
   * @returns The formatting protocol for the current workbook or null if not available
   */
  public getWorkbookFormattingProtocol(workbookId?: string): FormattingProtocol | null {
    const id = workbookId || this.currentWorkbookId;
    
    if (!id || !this.hasWorkbookFormatting(id)) {
      return null;
    }
    
    return this.workbookFormattingProtocols.get(id)!.formattingProtocol;
  }
  
  public hasWorkbookFormatting(workbookId: string): boolean {
    return this.workbookFormattingProtocols.has(workbookId);
  }
  /**
   * Gets the formatting metadata for the current workbook
   * @returns The formatting metadata for the current workbook or null if not available
   */
  public getWorkbookFormattingMetadata(workbookId?: string): WorkbookFormattingMetadata | null {
    const id = workbookId || this.currentWorkbookId;
    
    if (!id || !this.hasWorkbookFormatting(id)) {
      return null;
    }
    
    return this.workbookFormattingProtocols.get(id)!.formattingMetadata;
  }
  
  /**
   * Registers a workbook change to invalidate cached formatting
   * @param workbookId The ID of the workbook that changed
   */
  public registerWorkbookChange(workbookId: string): void {
    // Remove the workbook's formatting from the cache to force a re-analysis
    if (this.hasWorkbookFormatting(workbookId)) {
      console.log(`Invalidating formatting protocol for workbook: ${workbookId}`);
      this.workbookFormattingProtocols.delete(workbookId);
      
      // If this is the current workbook, also clear the legacy cache
      if (workbookId === this.currentWorkbookId) {
        this.cachedFormattingProtocol = null;
        this.cachedFormattingMetadata = null;
        this.lastAnalysisTimestamp = '';
      }
    }
  }
  
  /**
   * Refreshes images for a specific sheet without re-analyzing the formatting protocol
   * @param workbookId The ID of the workbook containing the sheet
   * @param sheetName The name of the sheet to refresh
   * @returns Promise that resolves when the sheet images have been refreshed
   */
  public async refreshSheetImages(workbookId: string, sheetName: string): Promise<void> {
    console.log(`Refreshing images for sheet: ${sheetName} in workbook: ${workbookId}`);
    
    try {
      // Find any cached analysis results for this workbook
      const analysisResults: MultimodalAnalysisResult[] = [];
      this.analysisResults.forEach((result) => {
        // Check if this result might be for the current workbook
        // We don't have a direct workbookId in the results, so we'll refresh all results
        // that include this sheet
        const includesSheet = result.metadata.options?.sheets?.includes(sheetName);
        if (includesSheet) {
          analysisResults.push(result);
        }
      });
      
      if (analysisResults.length === 0) {
        console.log(`No cached analysis results found for sheet: ${sheetName}`);
        return;
      }
      
      // Refresh the images for this specific sheet
      const refreshedImages = await performMultimodalAnalysis(this.apiEndpoint, {
        sheets: [sheetName]
      });
      
      // Update the cached results with the new images
      for (const result of analysisResults) {
        // Replace the images in the result
        result.images = refreshedImages;
        result.metadata.timestamp = new Date().toISOString();
      }
      
      console.log(`Successfully refreshed images for sheet: ${sheetName}`);
    } catch (error) {
      console.error(`Error refreshing images for sheet: ${sheetName}:`, error);
    }
  }

  /**
   * Analyzes the active workbook using multimodal techniques
   * @param options Optional configuration for the analysis
   * @returns Promise with analysis result including images
   */
  public async analyzeActiveWorkbook(options?: MultimodalAnalysisOptions): Promise<MultimodalAnalysisResult> {
    try {
      // Get workbook images through the utility function
      const images = await performMultimodalAnalysis(this.apiEndpoint, options);
      
      // Create result object
      const result: MultimodalAnalysisResult = {
        images,
        metadata: {
          timestamp: new Date().toISOString(),
          sheetCount: await this.getWorksheetCount(),
          options: options
        }
      };
      
      // Store result for future reference
      const analysisId = `analysis_${Date.now()}`;
      this.analysisResults.set(analysisId, result);
      
      return result;
    } catch (error) {
      console.error('Error in multimodal analysis service:', error);
      throw error;
    }
  }
  
  // Flag to track if a formatting analysis is currently in progress
  private isFormattingAnalysisInProgress = false;
  
  // Promise for the current formatting analysis
  private currentFormattingAnalysisPromise: Promise<FormattingProtocol> | null = null;
  
  /**
   * Analyzes the formatting protocol of the active workbook
   * @param waitForCompletion Whether to wait for the full analysis to complete (default: false)
   * @returns Promise with the formatting protocol analysis
   */
  public async analyzeFormattingProtocol(waitForCompletion: boolean = false): Promise<FormattingProtocol> {
    try {
      console.log(`
=======================================================
🔍 ANALYZE FORMATTING PROTOCOL (waitForCompletion=${waitForCompletion})
=======================================================`);
      
      // Check if we have a cached protocol for the current workbook
      if (this.currentWorkbookId && this.hasWorkbookFormatting(this.currentWorkbookId)) {
        console.log(`📋 Using cached formatting protocol for workbook: ${this.currentWorkbookId}`);
        const cachedProtocol = this.workbookFormattingProtocols.get(this.currentWorkbookId)!.formattingProtocol;
        console.log(`   Cached protocol has ${Object.keys(cachedProtocol).length} top-level categories`);
        return cachedProtocol;
      }
      
      // If waitForCompletion is true, perform the full analysis and wait for it to complete
      if (waitForCompletion) {
        // Add a clearer log message to show we're in blocking mode
        console.log(`
=======================================================
🕐 BLOCKING MODE: Performing FULL formatting analysis and waiting for completion
=======================================================`);
        
        // If there's already an analysis in progress, wait for it to complete
        if (this.isFormattingAnalysisInProgress && this.currentFormattingAnalysisPromise) {
          console.log('🕐 Waiting for in-progress formatting analysis to complete');
          try {
            // Wait for the existing promise
            const result = await this.currentFormattingAnalysisPromise;
            console.log(`✅ Successfully obtained result from in-progress analysis with ${Object.keys(result).length} categories`);
            
            // Store the result in the cache
            if (this.currentWorkbookId) {
              // Make sure we store the result in the cache
              this.workbookFormattingProtocols.set(this.currentWorkbookId, {
                formattingProtocol: result,
                formattingMetadata: this.cachedFormattingMetadata || this.getDefaultFormattingMetadata(),
                timestamp: new Date().toISOString()
              });
              
              console.log(`✅ Cached formatting protocol for workbook: ${this.currentWorkbookId}`);
            }
            
            return result;
          } catch (waitError) {
            console.error('❌ Error waiting for in-progress formatting analysis:', waitError);
            console.error(`Error stack trace: ${waitError.stack}`);
            throw waitError;
          }
        }
        
        // Otherwise, perform the analysis directly
        console.log('🕐 No analysis in progress, performing direct analysis');
        try {
          // Set the flag to indicate analysis is in progress
          this.isFormattingAnalysisInProgress = true;
          
          // Perform the analysis
          const result = await this.performFormattingAnalysis();
          
          // Store the result in the cache
          if (this.currentWorkbookId) {
            this.workbookFormattingProtocols.set(this.currentWorkbookId, {
              formattingProtocol: result,
              formattingMetadata: this.cachedFormattingMetadata || this.getDefaultFormattingMetadata(),
              timestamp: new Date().toISOString()
            });
            
            console.log(`✅ Cached formatting protocol for workbook: ${this.currentWorkbookId}`);
          }
          
          console.log(`
=======================================================
✅ BLOCKING MODE: Direct analysis completed with ${Object.keys(result).length} categories
=======================================================`);
          
          // Reset the analysis flag
          this.isFormattingAnalysisInProgress = false;
          
          return result;
        } catch (analysisError) {
          console.error('❌ Error in direct formatting analysis:', analysisError);
          console.error(`Error stack trace: ${analysisError.stack}`);
          
          // Reset the analysis flag
          this.isFormattingAnalysisInProgress = false;
          
          throw analysisError;
        }
      }
      
      // If waitForCompletion is false (default behavior), use the non-blocking approach
      // If there's already an analysis in progress, return a basic protocol immediately
      if (this.isFormattingAnalysisInProgress) {
        console.log('💭 Formatting analysis already in progress, returning basic protocol');
        return this.getBasicFormattingProtocol();
      }
      
      // Start the analysis in the background and return a basic protocol immediately
      console.log('💭 Starting formatting analysis in background');
      this.startFormattingAnalysisInBackground();
      
      // Return a basic protocol to avoid blocking the main flow
      console.log('💭 Returning basic protocol while analysis runs in background');
      return this.getBasicFormattingProtocol();
    } catch (error) {
      console.error('❌ Error in analyzeFormattingProtocol:', error);
      console.error(`Error stack trace: ${error.stack}`);
      
      // Check if this is specifically an Excel Image API error
      if (error.message && (
          error.message.includes('Excel Image API') ||
          error.message.includes('localhost:8080')
      )) {
        console.error(`
=======================================================
❌ FORMATTING PROTOCOL ANALYSIS FAILED: EXCEL IMAGE API ERROR
=======================================================
The formatting protocol analysis requires the Excel Image API server to be running.

Please make sure the Excel Image API server is running at http://localhost:8080.
This is required for formatting protocol analysis to work properly.

Using default formatting protocol as a fallback.
=======================================================`);
      } else {
        console.error(`
=======================================================
❌ FORMATTING PROTOCOL ANALYSIS FAILED: ${error.message}
=======================================================
Using default formatting protocol as a fallback.
=======================================================`);
      }
      
      return this.getBasicFormattingProtocol();
    }
  }
  
  /**
   * Starts the formatting analysis process in the background
   * This method doesn't block and allows the main flow to continue
   */
  private startFormattingAnalysisInBackground(): void {
    // If there's already an analysis in progress, don't start another one
    if (this.isFormattingAnalysisInProgress) {
      return;
    }
    
    // Set the flag to indicate that an analysis is in progress
    this.isFormattingAnalysisInProgress = true;
    
    // Start the analysis in the background
    this.currentFormattingAnalysisPromise = this.performFormattingAnalysis();
    
    // When the analysis completes, update the cache and reset the flag
    this.currentFormattingAnalysisPromise
      .then(formattingProtocol => {
        // Store the results in the cache
        if (this.currentWorkbookId) {
          this.workbookFormattingProtocols.set(this.currentWorkbookId, {
            formattingProtocol,
            formattingMetadata: this.cachedFormattingMetadata || this.getDefaultFormattingMetadata(),
            timestamp: new Date().toISOString()
          });
        }
        
        // Update the legacy cache
        this.cachedFormattingProtocol = formattingProtocol;
        this.lastAnalysisTimestamp = new Date().toISOString();
        
        console.log('Background formatting analysis completed successfully');
      })
      .catch(error => {
        console.error('Error in background formatting analysis:', error);
      })
      .finally(() => {
        // Reset the flag and promise
        this.isFormattingAnalysisInProgress = false;
        this.currentFormattingAnalysisPromise = null;
      });
  }
  
  /**
   * Performs the actual formatting analysis
   * This method is called by startFormattingAnalysisInBackground and runs asynchronously
   * @returns Promise with the formatting protocol analysis
   */
  private async performFormattingAnalysis(): Promise<FormattingProtocol> {
    try {
      console.log(`
=======================================================
🔍 FULL FORMATTING PROTOCOL ANALYSIS FOR WORKBOOK: ${this.currentWorkbookId}
=======================================================`);
      
      // Step 1: Extract formatting metadata (this doesn't require the API)
      console.log('📊 STEP 1: Extracting formatting metadata from workbook');
      const startMetadataTime = Date.now();
      const formattingMetadata = await this.formattingProtocolAnalyzer.extractFormattingMetadata();
      const metadataTime = Date.now() - startMetadataTime;
      console.log(`✅ Successfully extracted formatting metadata in ${metadataTime}ms`);
      console.log(`   Metadata contains information for ${formattingMetadata.sheets.length} sheets`);
      
      // Store the metadata in the cache immediately
      this.cachedFormattingMetadata = formattingMetadata;
      
      // Step 2: Get workbook images for all sheets
      console.log('🖼️ STEP 2: Converting worksheets to images');
      const worksheetNames = await this.getAllWorksheetNames();
      console.log(`   Requesting images for all ${worksheetNames.length} sheets in workbook`);
      
      // Request images for all sheets explicitly
      let images: string[] = [];
      const startImageTime = Date.now();
      try {
        // Use a timeout to prevent this from hanging indefinitely
        console.log('   Calling Excel Image API to convert worksheets to PNG images...');
        const imagePromise = performMultimodalAnalysis(this.apiEndpoint, {
          sheets: worksheetNames
        });
        
        // Set a timeout of 30 seconds for the image retrieval
        const timeoutPromise = new Promise<string[]>((_, reject) => {
          setTimeout(() => reject(new Error('Image retrieval timed out')), 30000);
        });
        
        // Race the image retrieval against the timeout
        images = await Promise.race([imagePromise, timeoutPromise]);
        const imageTime = Date.now() - startImageTime;
        console.log(`✅ Successfully retrieved ${images.length} images in ${imageTime}ms`);
      } catch (imageError) {
        console.warn('❌ Error retrieving workbook images:', imageError);
        // Continue with analysis using just the metadata
        images = [];
      }
      
      // Step 3: Analyze formatting protocol using the LLM
      console.log('🧠 STEP 3: Analyzing formatting with Claude LLM');
      console.log('   Sending images and metadata to Anthropic Claude for analysis...');
      // Even if we couldn't get images, we can still analyze the metadata
      const startLlmTime = Date.now();
      const formattingProtocol = await this.formattingProtocolAnalyzer.analyzeFormattingProtocol();
      const llmTime = Date.now() - startLlmTime;
      console.log(`✅ LLM analysis completed in ${llmTime}ms`);
      
      const totalTime = Date.now() - startMetadataTime;
      console.log(`
=======================================================
✅ FORMATTING PROTOCOL ANALYSIS COMPLETE: ${Math.round(totalTime/1000)}s
=======================================================
`);
      
      return formattingProtocol;
    } catch (error) {
      console.error('❌ Error in performFormattingAnalysis:', error);
      return this.getBasicFormattingProtocol();
    }
  }
  
  /**
   * Returns a basic formatting protocol to use when no cached protocol is available
   * or when an error occurs during analysis
   * @returns A basic formatting protocol
   */
  private getBasicFormattingProtocol(): FormattingProtocol {
    return {
      colorCoding: {
        inputs: '#f5f5f5',
        calculations: '#ffffff',
        headers: '#d0d0d0',
        totals: '#e0e0e0',
        custom: {}
      },
      numberFormatting: {
        currency: '$#,##0.00',
        percentage: '0.00%',
        date: 'mm/dd/yyyy',
        custom: {}
      },
      borderStyles: {
        tables: 'thin solid black',
        totals: 'medium solid black',
        custom: {}
      },
      fontUsage: {
        headers: {
          bold: true
        },
        body: {
          name: 'Arial'
        }
      },
      tableFormatting: {
        headerRow: {
          fontBold: true
        }
      },
      scheduleFormatting: {},
      workbookStructure: {},
      scenarioFormatting: {},
      chartFormatting: {
        chartTypes: {
          preferred: ['line', 'column', 'pie']
        },
        title: {
          hasTitle: true
        },
        legend: {
          position: 'right'
        }
      }
    };
  }
  
  /**
   * Returns default formatting metadata to use when no cached metadata is available
   * @returns A basic WorkbookFormattingMetadata object with default values
   */
  private getDefaultFormattingMetadata(): WorkbookFormattingMetadata {
    return {
      themeColors: {
        background1: '#FFFFFF',
        background2: '#F2F2F2',
        text1: '#000000',
        text2: '#666666',
        accent1: '#4472C4',
        accent2: '#ED7D31',
        accent3: '#A5A5A5',
        accent4: '#FFC000',
        accent5: '#5B9BD5',
        accent6: '#70AD47',
        hyperlink: '#0563C1',
        followedHyperlink: '#954F72'
      },
      sheets: []
    };
  }
  
  /**
   * Checks if the cached formatting protocol is still valid
   * @returns True if the protocol is valid, false otherwise
   */
  private isFormattingProtocolValid(): boolean {
    // If no cached protocol, it's not valid
    if (!this.cachedFormattingProtocol || !this.lastAnalysisTimestamp) {
      return false;
    }
    
    // Check if the protocol is less than 1 hour old
    const now = new Date();
    const lastAnalysis = new Date(this.lastAnalysisTimestamp);
    const hourInMs = 60 * 60 * 1000;
    
    return (now.getTime() - lastAnalysis.getTime()) < hourInMs;
  }
  
  /**
   * Checks if a workbook's formatting protocol is still valid
   * @param workbookId The ID of the workbook to check
   * @returns True if the protocol is valid, false otherwise
   */
  private isWorkbookFormattingValid(workbookId: string): boolean {
    // If no formatting for this workbook, it's not valid
    if (!this.hasWorkbookFormatting(workbookId)) {
      return false;
    }
    
    // Check if the protocol is less than 1 hour old
    const now = new Date();
    const workbookData = this.workbookFormattingProtocols.get(workbookId)!;
    const lastAnalysis = new Date(workbookData.timestamp);
    const hourInMs = 60 * 60 * 1000;
    
    return (now.getTime() - lastAnalysis.getTime()) < hourInMs;
  }
  
  /**
   * Gets the cached formatting protocol if available
   * @returns The cached formatting protocol or null if not available
   */
  public getCachedFormattingProtocol(): FormattingProtocol | null {
    // First try to get the formatting protocol for the current workbook
    if (this.currentWorkbookId && this.isWorkbookFormattingValid(this.currentWorkbookId)) {
      return this.workbookFormattingProtocols.get(this.currentWorkbookId)!.formattingProtocol;
    }
    
    // Fall back to the legacy cache for backward compatibility
    return this.isFormattingProtocolValid() ? this.cachedFormattingProtocol : null;
  }
  
  /**
   * Gets the cached formatting metadata if available
   * @returns The cached formatting metadata or null if not available
   */
  public getCachedFormattingMetadata(): WorkbookFormattingMetadata | null {
    // First try to get the formatting metadata for the current workbook
    if (this.currentWorkbookId && this.isWorkbookFormattingValid(this.currentWorkbookId)) {
      return this.workbookFormattingProtocols.get(this.currentWorkbookId)!.formattingMetadata;
    }
    
    // Fall back to the legacy cache for backward compatibility
    return this.isFormattingProtocolValid() ? this.cachedFormattingMetadata : null;
  }
  
  /**
   * Invalidates the cached formatting protocol
   */
  public invalidateFormattingProtocol(): void {
    // If we have a current workbook, invalidate its formatting
    if (this.currentWorkbookId) {
      this.registerWorkbookChange(this.currentWorkbookId);
    }
    
    // Also invalidate the legacy cache for backward compatibility
    this.cachedFormattingProtocol = null;
    this.cachedFormattingMetadata = null;
    this.lastAnalysisTimestamp = '';
    console.log('Formatting protocol cache invalidated');
  }
  
  /**
   * Analyzes specific ranges in the active workbook
   * @param ranges Array of sheet names and ranges to analyze
   * @returns Promise with analysis result including images
   */
  public async analyzeRanges(ranges: Array<{ sheetName: string; range: string }>): Promise<MultimodalAnalysisResult> {
    return this.analyzeActiveWorkbook({ ranges });
  }
  
  /**
   * Analyzes specific sheets in the active workbook
   * @param sheets Array of sheet names to analyze
   * @returns Promise with analysis result including images
   */
  public async analyzeSheets(sheets: string[]): Promise<MultimodalAnalysisResult> {
    return this.analyzeActiveWorkbook({ sheets });
  }
  
  /**
   * Analyzes charts in the active workbook
   * @param charts Array of chart specifications to analyze
   * @returns Promise with analysis result including images
   */
  public async analyzeCharts(charts: Array<{ sheetName: string; chartIndex: number }>): Promise<MultimodalAnalysisResult> {
    return this.analyzeActiveWorkbook({ charts });
  }

  /**
   * Gets the number of worksheets in the active workbook
   */
  private async getWorksheetCount(): Promise<number> {
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items');
      await context.sync();
      return worksheets.items.length;
    });
  }
  
  /**
   * Gets the names of all worksheets in the active workbook
   * @returns Promise with array of worksheet names
   */
  private async getAllWorksheetNames(): Promise<string[]> {
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items/name');
      await context.sync();
      
      return worksheets.items.map(sheet => sheet.name);
    });
  }

  /**
   * Retrieves a previously stored analysis result
   * @param analysisId The ID of the analysis to retrieve
   */
  public getAnalysisResult(analysisId: string): MultimodalAnalysisResult | undefined {
    return this.analysisResults.get(analysisId);
  }

  /**
   * Gets all stored analysis results
   */
  public getAllAnalysisResults(): MultimodalAnalysisResult[] {
    return Array.from(this.analysisResults.values());
  }


  /**
   * Gets the current workbook ID
   */
  public getCurrentWorkbookId(): string | undefined {
    return this.currentWorkbookId;
  }
}

// Create singleton instance (without Anthropic service for now)
// The actual Anthropic service should be injected when the application starts
export const multimodalAnalysisService = new MultimodalAnalysisService();

// Add new persistence methods to MultimodalAnalysisService
Object.assign(MultimodalAnalysisService.prototype, {
  /**
   * Saves the formatting protocol for a workbook to Office.context.document.settings
   * @param workbookId The ID of the workbook
   * @param data The formatting data to save
   */
  saveFormattingProtocolToSettings(workbookId: string, data: any): void {
    try {
      if (Office?.context?.document?.settings) {
        const settingsKey = `formatting_protocol_${workbookId}`;
        
        // We need to stringify the data to store it in settings
        const serializedData = JSON.stringify(data);
        
        // Save to document settings
        Office.context.document.settings.set(settingsKey, serializedData);
        
        // Persist the settings
        Office.context.document.settings.saveAsync(() => {
          console.log(`✅ Successfully saved formatting protocol for workbook ${workbookId} to document settings`);
        });
      } else {
        console.warn('⚠️ Cannot save formatting protocol: Office context or settings not available');
      }
    } catch (error) {
      console.error('❌ Error saving formatting protocol to settings:', error);
    }
  },
  
  /**
   * Loads the formatting protocol for a workbook from Office.context.document.settings
   * @param workbookId The ID of the workbook
   * @returns The formatting data, or undefined if not found
   */
  loadFormattingProtocolFromSettings(workbookId: string): {
    formattingProtocol: FormattingProtocol;
    formattingMetadata: WorkbookFormattingMetadata;
    timestamp: string;
  } | undefined {
    try {
      if (Office?.context?.document?.settings) {
        const settingsKey = `formatting_protocol_${workbookId}`;
        
        // Get the serialized data from settings
        const serializedData = Office.context.document.settings.get(settingsKey);
        
        if (serializedData) {
          // Parse the serialized data
          const parsedData = JSON.parse(serializedData);
          console.log(`✅ Successfully loaded formatting protocol for workbook ${workbookId} from document settings`);
          return parsedData;
        }
      } else {
        console.warn('⚠️ Cannot load formatting protocol: Office context or settings not available');
      }
    } catch (error) {
      console.error('❌ Error loading formatting protocol from settings:', error);
    }
    
    return undefined;
  },
  
  /**
   * Loads all cached formatting protocols from settings
   */
  loadCachedFormattingProtocols(): void {
    try {
      if (Office?.context?.document?.settings) {
        // Try to get all keys that start with 'formatting_protocol_'
        const allKeys = Object.keys(Office.context.document.settings);
        const protocolKeys = allKeys.filter(key => key.startsWith('formatting_protocol_'));
        
        for (const key of protocolKeys) {
          // Extract the workbook ID from the key
          const workbookId = key.replace('formatting_protocol_', '');
          
          // Load the formatting protocol
          const formattingData = this.loadFormattingProtocolFromSettings(workbookId);
          
          if (formattingData) {
            // Store in memory
            this.workbookFormattingProtocols.set(workbookId, formattingData);
          }
        }
        
        console.log(`✅ Loaded ${this.workbookFormattingProtocols.size} formatting protocols from settings`);
      }
    } catch (error) {
      console.error('❌ Error loading cached formatting protocols:', error);
    }
  }
});

/**
 * Initialize the multimodal analysis service with the Anthropic service
 * @param anthropicService The Anthropic service instance
 * @param workbookId Optional current workbook ID
 */
export function initializeMultimodalAnalysisService(
  anthropicService: ClientAnthropicService,
  workbookId?: string
): void {
  // Create a new service with the provided Anthropic service
  const newService = new MultimodalAnalysisService(
    DEFAULT_IMAGE_CONVERSION_ENDPOINT,
    anthropicService
  );
  
  // Replace the singleton instance with the new service
  Object.assign(multimodalAnalysisService, newService);
  
  // If a workbook ID was provided, set it as the current workbook
  if (workbookId) {
    // We don't await this since it's an initialization function
    // The formatting will be analyzed when needed
    multimodalAnalysisService.setWorkbookAndEnsureFormatting(workbookId);
  }
  
  console.log('Multimodal analysis service initialized successfully');
}
