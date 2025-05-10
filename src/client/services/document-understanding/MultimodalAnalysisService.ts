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
   * @returns Promise that resolves when the workbook's formatting has been analyzed
   */
  public async setWorkbookAndEnsureFormatting(workbookId: string): Promise<void> {
    // If this is a different workbook than the current one
    if (this.currentWorkbookId !== workbookId) {
      console.log(`Setting current workbook to: ${workbookId}`);
      this.currentWorkbookId = workbookId;
      
      // Check if we've already analyzed this workbook
      if (!this.hasWorkbookFormatting(workbookId)) {
        console.log(`Analyzing formatting protocol for workbook: ${workbookId}`);
        
        try {
          // Analyze the formatting protocol for this workbook
          const formattingProtocol = await this.analyzeFormattingProtocol();
          
          // Store the formatting protocol for this workbook
          this.workbookFormattingProtocols.set(workbookId, {
            formattingProtocol,
            formattingMetadata: this.cachedFormattingMetadata!,
            timestamp: new Date().toISOString()
          });
          
          console.log(`Formatting protocol analysis complete for workbook: ${workbookId}`);
        } catch (error) {
          console.error(`Error analyzing formatting protocol for workbook ${workbookId}:`, error);
          throw error;
        }
      } else {
        console.log(`Using cached formatting protocol for workbook: ${workbookId}`);
        
        // Update the legacy cache variables for backward compatibility
        const workbookData = this.workbookFormattingProtocols.get(workbookId)!;
        this.cachedFormattingProtocol = workbookData.formattingProtocol;
        this.cachedFormattingMetadata = workbookData.formattingMetadata;
        this.lastAnalysisTimestamp = workbookData.timestamp;
      }
    }
  }
  
  /**
   * Checks if a workbook's formatting has been analyzed
   * @param workbookId The ID of the workbook to check
   * @returns True if the workbook's formatting has been analyzed
   */
  private hasWorkbookFormatting(workbookId: string): boolean {
    return this.workbookFormattingProtocols.has(workbookId);
  }
  
  /**
   * Gets the complete formatting data for a specific workbook
   * @param workbookId The ID of the workbook
   * @returns The complete formatting data for the workbook, or undefined if not available
   */
  public getWorkbookFormattingData(workbookId: string): {
    formattingProtocol: FormattingProtocol;
    formattingMetadata: WorkbookFormattingMetadata;
    timestamp: string;
  } | undefined {
    return this.workbookFormattingProtocols.get(workbookId);
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
   * Analyzes the formatting protocol of the active workbook without blocking the main flow
   * @returns Promise with the formatting protocol analysis
   */
  public async analyzeFormattingProtocol(): Promise<FormattingProtocol> {
    try {
      // Check if we have a cached protocol for the current workbook
      if (this.currentWorkbookId && this.hasWorkbookFormatting(this.currentWorkbookId)) {
        console.log(`Using cached formatting protocol for workbook: ${this.currentWorkbookId}`);
        return this.workbookFormattingProtocols.get(this.currentWorkbookId)!.formattingProtocol;
      }
      
      // If there's already an analysis in progress, return a basic protocol immediately
      // This prevents blocking the main flow while still allowing the analysis to complete in the background
      if (this.isFormattingAnalysisInProgress) {
        console.log('Formatting analysis already in progress, returning basic protocol');
        return this.getBasicFormattingProtocol();
      }
      
      // Start the analysis in the background and return a basic protocol immediately
      this.startFormattingAnalysisInBackground();
      
      // Return a basic protocol to avoid blocking the main flow
      return this.getBasicFormattingProtocol();
    } catch (error) {
      console.error('Error in analyzeFormattingProtocol:', error);
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
      console.log(`Analyzing formatting protocol for workbook: ${this.currentWorkbookId}`);
      
      // Step 1: Extract formatting metadata (this doesn't require the API)
      const formattingMetadata = await this.formattingProtocolAnalyzer.extractFormattingMetadata();
      console.log('Successfully extracted formatting metadata');
      
      // Store the metadata in the cache immediately
      this.cachedFormattingMetadata = formattingMetadata;
      
      // Step 2: Get workbook images for all sheets
      const worksheetNames = await this.getAllWorksheetNames();
      console.log(`Requesting images for all ${worksheetNames.length} sheets in workbook`);
      
      // Request images for all sheets explicitly
      let images: string[] = [];
      try {
        // Use a timeout to prevent this from hanging indefinitely
        const imagePromise = performMultimodalAnalysis(this.apiEndpoint, {
          sheets: worksheetNames
        });
        
        // Set a timeout of 30 seconds for the image retrieval
        const timeoutPromise = new Promise<string[]>((_, reject) => {
          setTimeout(() => reject(new Error('Image retrieval timed out')), 30000);
        });
        
        // Race the image retrieval against the timeout
        images = await Promise.race([imagePromise, timeoutPromise]);
        console.log(`Successfully retrieved ${images.length} images for analysis`);
      } catch (imageError) {
        console.warn('Error retrieving workbook images:', imageError);
        // Continue with analysis using just the metadata
        images = [];
      }
      
      // Step 3: Analyze formatting protocol using the LLM
      // Even if we couldn't get images, we can still analyze the metadata
      const formattingProtocol = await this.formattingProtocolAnalyzer.analyzeFormattingProtocol();
      
      return formattingProtocol;
    } catch (error) {
      console.error('Error in performFormattingAnalysis:', error);
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
}

// Create singleton instance (without Anthropic service for now)
// The actual Anthropic service should be injected when the application starts
export const multimodalAnalysisService = new MultimodalAnalysisService();

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
