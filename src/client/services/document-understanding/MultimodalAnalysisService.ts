/**
 * Service for multimodal analysis of Excel workbooks
 * Integrates with LLMs for formatting analysis
 */
import { performMultimodalAnalysis } from './WorkbookUtils';
import { FormattingProtocolAnalyzer, FormattingProtocol, WorkbookFormattingMetadata } from './FormattingProtocolAnalyzer';
import { ClientAnthropicService } from '../ClientAnthropicService';

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
  
  /**
   * Analyzes the formatting protocol of the active workbook using LLM
   * @returns Promise with the formatting protocol analysis
   */
  public async analyzeFormattingProtocol(): Promise<FormattingProtocol> {
    try {
      // Check if we have a cached protocol that's still valid
      if (this.cachedFormattingProtocol && this.isFormattingProtocolValid()) {
        console.log('Using cached formatting protocol');
        return this.cachedFormattingProtocol;
      }
      
      // Ensure we have the formatting protocol analyzer
      if (!this.formattingProtocolAnalyzer) {
        throw new Error('Formatting protocol analyzer not initialized. Please provide an Anthropic service when creating the MultimodalAnalysisService.');
      }
      
      // Step 1: Get workbook images for all sheets
      const images = await performMultimodalAnalysis(this.apiEndpoint);
      
      // Step 2: Extract formatting metadata
      const formattingMetadata = await this.formattingProtocolAnalyzer.extractFormattingMetadata();
      
      // Step 3: Analyze the formatting protocol using LLM
      const formattingProtocol = await this.formattingProtocolAnalyzer.analyzeFormattingProtocol(
        images,
        formattingMetadata
      );
      
      // Cache the results
      this.cachedFormattingProtocol = formattingProtocol;
      this.cachedFormattingMetadata = formattingMetadata;
      this.lastAnalysisTimestamp = new Date().toISOString();
      
      // Store in analysis results
      const analysisId = `formatting_protocol_${Date.now()}`;
      this.analysisResults.set(analysisId, {
        images,
        formattingAnalysis: formattingProtocol,
        formattingMetadata: formattingMetadata,
        metadata: {
          timestamp: this.lastAnalysisTimestamp,
          sheetCount: formattingMetadata.sheets.length
        }
      });
      
      return formattingProtocol;
    } catch (error) {
      console.error('Error analyzing formatting protocol:', error);
      throw error;
    }
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
   * Gets the cached formatting protocol if available
   * @returns The cached formatting protocol or null if not available
   */
  public getCachedFormattingProtocol(): FormattingProtocol | null {
    return this.isFormattingProtocolValid() ? this.cachedFormattingProtocol : null;
  }
  
  /**
   * Gets the cached formatting metadata if available
   * @returns The cached formatting metadata or null if not available
   */
  public getCachedFormattingMetadata(): WorkbookFormattingMetadata | null {
    return this.isFormattingProtocolValid() ? this.cachedFormattingMetadata : null;
  }
  
  /**
   * Invalidates the cached formatting protocol
   */
  public invalidateFormattingProtocol(): void {
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
      worksheets.load("items");
      await context.sync();
      return worksheets.items.length;
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
 */
export function initializeMultimodalAnalysisService(anthropicService: ClientAnthropicService): void {
  // Create a new instance with the Anthropic service
  const newService = new MultimodalAnalysisService(
    DEFAULT_IMAGE_CONVERSION_ENDPOINT,
    anthropicService
  );
  
  // Copy properties from the new service to the singleton
  Object.assign(multimodalAnalysisService, newService);
  
  console.log('Multimodal analysis service initialized with Anthropic service');
}
