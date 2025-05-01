import { MetadataChunk, WorkbookMetrics } from '../models/CommandModels';
import { SpreadsheetChunkCompressor } from './SpreadsheetChunkCompressor';
import { RangeDependencyAnalyzer } from './RangeDependencyAnalyzer';

/**
 * Cache for workbook metadata chunks with invalidation and dependency tracking
 */
export class WorkbookMetadataCache {
  private chunks: Map<string, MetadataChunk> = new Map();
  private compressor: SpreadsheetChunkCompressor;
  private dependencyAnalyzer: RangeDependencyAnalyzer;
  private workbookVersion: string = '';
  
  constructor() {
    this.compressor = new SpreadsheetChunkCompressor();
    this.dependencyAnalyzer = new RangeDependencyAnalyzer();
  }

  /**
   * Get a cached chunk by its ID, or null if not cached
   * @param chunkId The chunk ID
   * @returns The metadata chunk or null if not found
   */
  public getChunk(chunkId: string): MetadataChunk | null {
    return this.chunks.get(chunkId) || null;
  }

  /**
   * Store a chunk in the cache
   * @param chunk The metadata chunk to store
   */
  public setChunk(chunk: MetadataChunk): void {
    this.chunks.set(chunk.id, chunk);
    console.log(`%c Cached chunk: ${chunk.id}`, 'color: #3498db');
  }

  /**
   * Check if a chunk exists in the cache
   * @param chunkId The chunk ID
   * @returns True if the chunk exists in the cache
   */
  public hasChunk(chunkId: string): boolean {
    return this.chunks.has(chunkId);
  }

  /**
   * Get all cached chunks
   * @returns Array of all cached metadata chunks
   */
  public getAllChunks(): MetadataChunk[] {
    return Array.from(this.chunks.values());
  }

  /**
   * Get all sheet chunks from the cache
   * @returns Array of sheet chunks
   */
  public getAllSheetChunks(): MetadataChunk[] {
    return Array.from(this.chunks.values())
      .filter(chunk => chunk.type === 'sheet');
  }

  /**
   * Invalidate specific chunks in the cache
   * @param chunkIds Array of chunk IDs to invalidate
   */
  public invalidateChunks(chunkIds: string[]): void {
    if (!chunkIds || chunkIds.length === 0) {
      return;
    }

    console.log(`%c Invalidating ${chunkIds.length} chunks`, 'color: #e74c3c');
    
    // Track affected chunks including dependencies and dependents
    const affectedChunks = new Set<string>();
    
    // Add explicitly invalidated chunks
    chunkIds.forEach(id => {
      affectedChunks.add(id);
    });
    
    // Find chunks that depend on the invalidated chunks (dependents)
    const dependents = this.dependencyAnalyzer.getTransitiveDependents(chunkIds);
    dependents.forEach(id => affectedChunks.add(id));
    
    // Delete all affected chunks from the cache
    affectedChunks.forEach(id => {
      this.chunks.delete(id);
    });
    
    console.log(`%c Invalidated chunks: ${Array.from(affectedChunks).join(', ')}`, 'color: #e74c3c');
  }
  
  /**
   * Invalidate all chunks for a specific sheet
   * @param sheetName The name of the sheet to invalidate
   */
  public invalidateChunksForSheet(sheetName: string): void {
    const sheetId = `Sheet:${sheetName}`;
    
    // First check if this sheet is in the cache
    if (!this.hasChunk(sheetId)) {
      console.log(`%c Sheet ${sheetName} not in cache, nothing to invalidate`, 'color: #95a5a6');
      return;
    }
    
    console.log(`%c Invalidating all chunks for sheet: ${sheetName}`, 'color: #e74c3c');
    
    // Invalidate the sheet chunk and any range chunks that belong to this sheet
    const chunksToInvalidate = Array.from(this.chunks.keys()).filter(id => {
      // Include the sheet itself
      if (id === sheetId) return true;
      
      // Include any ranges from this sheet (Range:SheetName!A1:B10 format)
      if (id.startsWith(`Range:${sheetName}!`)) return true;
      
      return false;
    });
    
    if (chunksToInvalidate.length > 0) {
      this.invalidateChunks(chunksToInvalidate);
    }
  }
  
  /**
   * Get or capture a sheet chunk by its name
   * @param sheetName The name of the sheet
   * @param forceRefresh Whether to force refresh the chunk
   * @param workbookStateManager Optional workbook state manager to capture the sheet if not cached
   * @returns The sheet chunk or null if it couldn't be captured
   */
  public async getOrCaptureSheet(
    sheetName: string, 
    forceRefresh: boolean = false,
    workbookStateManager?: any
  ): Promise<MetadataChunk | null> {
    const sheetId = `Sheet:${sheetName}`;
    
    // Return cached chunk if available and not forcing refresh
    if (!forceRefresh && this.hasChunk(sheetId)) {
      console.log(`%c Using cached sheet chunk for: ${sheetName}`, 'color: #27ae60');
      return this.getChunk(sheetId);
    }
    
    // If no workbook state manager provided, can't capture
    if (!workbookStateManager) {
      console.warn(`%c Cannot capture sheet ${sheetName}: No workbook state manager provided`, 'color: #e74c3c');
      return null;
    }
    
    console.log(`%c Capturing sheet chunk for: ${sheetName}`, 'color: #3498db');
    
    try {
      // Use the workbook state manager to capture the sheet
      return await Excel.run(async (context) => {
        // Get the worksheet
        const worksheet = context.workbook.worksheets.getItem(sheetName);
        
        // Capture the sheet state
        const sheetState = await workbookStateManager.captureSheetState(worksheet);
        
        // Compress the sheet into a chunk
        const chunk = workbookStateManager.getChunkCompressor().compressSheetToChunk(sheetState);
        
        // Add the chunk to the cache with dependency analysis
        this.addChunkWithDependencyAnalysis(chunk);
        
        return chunk;
      });
    } catch (error) {
      console.error(`%c Error capturing sheet ${sheetName}:`, 'color: #e74c3c', error);
      return null;
    }
  }

  /**
   * Invalidate all chunks in the cache
   */
  public invalidateAllChunks(): void {
    const chunkCount = this.chunks.size;
    console.log(`%c Invalidating all ${chunkCount} chunks`, 'color: #e74c3c');
    this.chunks.clear();
    this.dependencyAnalyzer.resetDependencyGraph();
    this.workbookVersion = '';
  }

  /**
   * Set the workbook version (e.g., a hash of workbook state)
   * @param version The workbook version string
   */
  public setWorkbookVersion(version: string): void {
    this.workbookVersion = version;
  }

  /**
   * Get the current workbook version
   * @returns The workbook version string
   */
  public getWorkbookVersion(): string {
    return this.workbookVersion;
  }

  /**
   * Calculate workbook metrics from currently cached chunks
   * @returns Aggregated workbook metrics
   */
  public calculateWorkbookMetrics(): WorkbookMetrics {
    // Use the compressor to calculate metrics from all cached chunks
    return this.compressor.calculateWorkbookMetrics(this.getAllChunks());
  }
  
  /**
   * Get the dependency analyzer
   * @returns The RangeDependencyAnalyzer instance
   */
  public getDependencyAnalyzer(): RangeDependencyAnalyzer {
    return this.dependencyAnalyzer;
  }
  
  /**
   * Add a new chunk to the cache and analyze its dependencies
   * @param chunk The metadata chunk to add
   */
  public addChunkWithDependencyAnalysis(chunk: MetadataChunk): void {
    // Store the chunk in the cache
    this.setChunk(chunk);
    
    // Analyze this chunk's dependencies
    this.dependencyAnalyzer.analyzeChunks([chunk]);
    
    // If this is a sheet chunk with formulas, analyze formula dependencies
    if (chunk.type === 'sheet' && chunk.payload && chunk.payload.formulas) {
      const sheetName = chunk.payload.name;
      const formulas = chunk.payload.formulas;
      
      // Analyze formulas to find inter-sheet references
      this.dependencyAnalyzer.analyzeFormulasInSheet(sheetName, formulas);
    }
  }
  
  /**
   * Get related chunks for a given query context
   * @param primaryChunkIds The IDs of the primary chunks for the query
   * @returns All chunks needed for the query context
   */
  public getRelatedChunks(primaryChunkIds: string[]): MetadataChunk[] {
    // Get the IDs of all chunks related to the primary ones
    const allChunkIds = new Set<string>(primaryChunkIds);
    
    // Add dependencies (chunks that the primary chunks depend on)
    const dependencies = this.dependencyAnalyzer.getTransitiveDependencies(primaryChunkIds);
    dependencies.forEach(id => allChunkIds.add(id));
    
    // Retrieve all chunks from the cache
    const result: MetadataChunk[] = [];
    
    allChunkIds.forEach(id => {
      const chunk = this.getChunk(id);
      if (chunk) {
        result.push(chunk);
      }
    });
    
    return result;
  }
}
