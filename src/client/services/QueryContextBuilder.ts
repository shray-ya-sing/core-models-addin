import { MetadataChunk, QueryContext, QueryType } from '../models/CommandModels';
import { ClientWorkbookStateManager } from './ClientWorkbookStateManager';
import { WorkbookMetadataCache } from './WorkbookMetadataCache';
import { ChunkLocatorService } from './ChunkLocatorService';

/**
 * Builds query contexts by selecting and assembling relevant chunks
 * for a given query or command
 */
export class QueryContextBuilder {
  private workbookStateManager: ClientWorkbookStateManager;
  private metadataCache: WorkbookMetadataCache;
  private chunkLocator: ChunkLocatorService;
  
  constructor(
    workbookStateManager: ClientWorkbookStateManager,
    metadataCache: WorkbookMetadataCache,
    chunkLocator: ChunkLocatorService
  ) {
    this.workbookStateManager = workbookStateManager;
    this.metadataCache = metadataCache;
    this.chunkLocator = chunkLocator;
  }
  
  /**
   * Set the chunk locator service
   * @param chunkLocator The chunk locator service
   */
  public setChunkLocator(chunkLocator: ChunkLocatorService): void {
    this.chunkLocator = chunkLocator;
    console.log('%c ChunkLocator service attached to QueryContextBuilder', 'color: #2ecc71');
  }

  /**
   * Build a query context for a specific query
   * In Phase 1, this will selectively include only relevant sheets and their dependencies
   * 
   * @param query The query text
   * @param queryType The type of query (question, command, etc.)
   * @param chatHistory The chat history
   * @param stepQuery The step query
   * @returns A query context with the relevant chunks
   */
  public async buildContextForQuery(
    queryType: QueryType,
    chatHistory: Array<{role: string, content: string}>,
    stepQuery: string
  ): Promise<QueryContext> {
    console.log(
      `%c Building query context for: "${stepQuery.substring(0, 50)}${stepQuery.length > 50 ? '...' : ''}"`,
      'queryType:', queryType,
      'background: #2c3e50; color: #ecf0f1; font-size: 12px; padding: 2px 5px;'
    );

    // First, ensure all sheets are captured and cached
    // This ensures we have complete dependency information
    await this.ensureAllSheetsCached(false);
    
    // Identify relevant sheets based on the step query
    const relevantSheetIds = await this.identifyRelevantSheets(stepQuery, chatHistory);
    
    // If we couldn't identify any specific sheets, include all sheets
    if (relevantSheetIds.length === 0) {
      console.log('%c No specific sheets identified, including all sheets', 'color: #e74c3c');
      return this.buildFullWorkbookContext();
    }
    
    // Get all related chunks (including dependencies)
    const chunks = this.metadataCache.getRelatedChunks(relevantSheetIds);
    
    // Get the active sheet name
    const activeSheet = await this.getActiveSheetName();
    
    // Build and return the context
    const context: QueryContext = {
      chunks,
      activeSheet,
      metrics: this.metadataCache.calculateWorkbookMetrics()
    };
    
    console.log(
      `%c Query context built with ${context.chunks.length} chunks out of ${this.metadataCache.getAllChunks().length} total chunks`,
      'background: #27ae60; color: #ecf0f1; font-size: 12px; padding: 2px 5px;'
    );
    
    // Log the included sheet names
    const includedSheetNames = chunks
      .filter(chunk => chunk.type === 'sheet')
      .map(chunk => chunk.payload.name)
      .join(', ');
    
    console.log(`%c Included sheets: ${includedSheetNames}`, 'color: #3498db');
    
    return context;
  }

  /**
   * Build a context that includes all sheets
   * Used as a fallback when specific sheets can't be identified
   * @returns A query context with all sheets
   */
  private async buildFullWorkbookContext(): Promise<QueryContext> {
    console.log('%c Building full workbook context', 'color: #9b59b6');
    
    // Get all available sheet chunks
    const chunks = this.metadataCache.getAllSheetChunks();
    const activeSheet = await this.workbookStateManager.getActiveSheetName() || '';
    
    // If we have no chunks but know we have sheets (from workbook capture),
    // create a minimal context with basic active sheet info
    if (chunks.length === 0) {
      console.warn('%c No valid sheet chunks available, creating minimal context', 'color: #e74c3c');
      
      // Try to get raw workbook state to create basic info
      try {
        const workbookState = await this.workbookStateManager.captureWorkbookState();
        const sheetNames = workbookState.sheets.map(s => s.name);
        
        console.log(`%c Found sheets: ${sheetNames.join(', ')}`, 'color: #f39c12');
        
        // Create minimal context with just the sheet names
        return {
          chunks: [],
          activeSheet,
          metrics: {
            totalSheets: sheetNames.length,
            totalCells: 0,  // We can't determine this without valid chunks
            totalFormulas: 0,
            totalTables: 0,
            totalCharts: 0
          }
        };
      } catch (error) {
        console.error('Failed to create even minimal context:', error);
      }
    }
    
    return {
      chunks,
      activeSheet,
      metrics: this.metadataCache.calculateWorkbookMetrics()
    };
  }

  /**
   * Ensure all sheets in the workbook are captured and cached
   * @param forceRefresh Whether to force refresh all chunks
   */
  private async ensureAllSheetsCached(forceRefresh: boolean): Promise<void> {
    // If we're not forcing a refresh and we have cached chunks, we're done
    if (!forceRefresh && this.metadataCache.getAllSheetChunks().length > 0) {
      console.log('%c Using cached sheet chunks', 'color: #3498db');
      return;
    }
    
    console.log('%c Capturing all sheets as chunks', 'color: #f39c12');
    
    try {
      // Capture the workbook state using the workbook manager
      const workbookState = await this.workbookStateManager.captureWorkbookState();
      
      // Track successful chunks for debugging
      let successCount = 0;
      let failCount = 0;
      
      // Process each sheet into a chunk and cache it with dependency analysis
      for (const sheet of workbookState.sheets) {
        try {
          const chunk = this.workbookStateManager.getChunkCompressor().compressSheetToChunk(sheet);
          this.metadataCache.addChunkWithDependencyAnalysis(chunk);
          console.log(`%c ✓ Cached and analyzed sheet: ${sheet.name}`, 'color: #2ecc71');
          successCount++;
        } catch (error) {
          console.error(`%c ✗ Error compressing sheet "${sheet.name}": ${error.message}`, 'color: #e74c3c');
          failCount++;
          
          // Create a simplified chunk with just the basic info to avoid completely missing this sheet
          try {
            // Create a basic chunk with minimal info
            const basicChunk = {
              id: `Sheet:${sheet.name}`,
              type: 'sheet' as 'sheet',
              etag: new Date().getTime().toString(), // Simple timestamp as etag
              payload: {
                name: sheet.name,
                summary: `Sheet ${sheet.name} (basic info only due to processing error)`,
                anchors: [],
                values: []
              },
              refs: [],
              lastCaptured: new Date()
            };
            
            this.metadataCache.setChunk(basicChunk);
            console.log(`%c ⚠ Created simplified chunk for sheet: ${sheet.name}`, 'color: #f39c12');
          } catch (fallbackError) {
            console.error(`%c Failed to create even a basic chunk for ${sheet.name}:`, 'color: #e74c3c', fallbackError);
          }
        }
      }
      
      console.log(`%c Sheet processing complete: ${successCount} succeeded, ${failCount} had errors`, 
                  successCount > 0 ? 'color: #2ecc71' : 'color: #e74c3c');
                  
      // If all sheets failed, throw an error so the caller knows
      if (successCount === 0 && failCount > 0) {
        throw new Error(`Failed to process any sheets successfully (${failCount} sheets had errors)`);
      }
    } catch (error) {
      console.error('%c Error capturing workbook state:', 'color: #e74c3c', error);
      throw error;
    }
  }

  /**
   * Identify the relevant sheets based on the query text
   * @param query The query text
   * @returns Array of sheet chunk IDs that are relevant to the query
   */
  private async identifyRelevantSheets(query: string, chatHistory: Array<{role: string, content: string}>): Promise<string[]> {
    // Debug logging to verify query value
    console.log(`%c QueryContextBuilder: Identifying sheets for query: "${query}"`, 'background: #e67e22; color: white; font-weight: bold; padding: 2px 5px;');
    console.log(`%c Query length: ${query.length}, First char code: ${query.charCodeAt(0)}, Last char code: ${query.charCodeAt(query.length-1)}`, 'color: #d35400;');
    
    // If we have a chunk locator service, use it for more advanced identification
    if (this.chunkLocator) {
      console.log('%c Using ChunkLocator service to identify relevant sheets', 'color: #3498db');
      
      try {
        // Update the chunk locator with the current active sheet
        const activeSheet = await this.getActiveSheetName();
        this.chunkLocator.setActiveSheet(activeSheet);
        
        // Use the chunk locator to find relevant chunks
        const locatorResult = await this.chunkLocator.locateChunks(query, chatHistory);
        
        console.log(`%c ChunkLocator identified ${locatorResult.chunkIds.length} relevant chunks`, 'color: #2ecc71');
        console.log(`%c Relevant sheets: ${locatorResult.details.sheets.join(', ')}`, 'color: #3498db');
        
        return locatorResult.chunkIds;
      } catch (error) {
        console.error('Error using ChunkLocator:', error);
      }
    }
    
    // Legacy implementation for backward compatibility
    // Enhanced keyword-based sheet identification with NLP-like features
    const allSheets = this.metadataCache.getAllSheetChunks();    
    return allSheets.map(sheet => sheet.id);
  }

  /**
   * Get the active sheet name
   * @returns Promise with the active sheet name
   */
  private async getActiveSheetName(): Promise<string> {
    return this.workbookStateManager.getActiveSheetName();
  }
  
  /**
   * Convert a QueryContext to a JSON string for the LLM
   * @param context The query context
   * @returns JSON string representing the workbook
   */
  public contextToJson(context: QueryContext): string {
    // For compatibility with the existing format,
    // convert our chunks back into the CompressedWorkbook format
    
    // Filter valid sheet chunks and extract their payloads
    const sheetPayloads = context.chunks
      .filter(chunk => chunk && chunk.type === 'sheet' && chunk.payload)
      .map(chunk => chunk.payload);
    
    // If no valid sheets were found, create a basic workbook structure
    // to avoid sending completely empty context to the LLM
    const sheets = sheetPayloads.length > 0 ? sheetPayloads : [
      {
        name: context.activeSheet || 'Sheet1',
        summary: 'Basic sheet information (limited data available)',
        anchors: [],
        values: []
      }
    ];
    
    // Also ensure we have valid metrics
    const metrics = context.metrics && typeof context.metrics === 'object' ? 
      context.metrics : 
      {
        totalSheets: sheets.length,
        totalCells: 0,
        totalFormulas: 0,
        totalTables: 0,
        totalCharts: 0
      };
    
    // Create the compressed workbook in the format expected by the LLM
    const compressedWorkbook = {
      sheets,
      activeSheet: context.activeSheet || (sheets[0] ? sheets[0].name : ''),
      metrics
    };
    
    // Add diagnostics to help debug issues
    if (process.env.NODE_ENV !== 'production') {
      compressedWorkbook['_diagnostic'] = {
        chunkCount: context.chunks.length,
        validSheetCount: sheetPayloads.length,
        timestamp: new Date().toISOString()
      };
    }
    
    return JSON.stringify(compressedWorkbook);
  }
}
