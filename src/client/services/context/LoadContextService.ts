import { ClientWorkbookStateManager } from "./ClientWorkbookStateManager";
import { ClientSpreadsheetCompressor } from "./ClientSpreadsheetCompressor";
import { QueryContextBuilder } from './QueryContextBuilder';
import { ChunkLocatorService } from './ChunkLocatorService';
import { ClientAnthropicService } from '../llm/ClientAnthropicService';
import { MistralClientService } from '../llm/MistralClientService';
import { EmbeddingService } from './EmbeddingService';
 
 /* --------------------------  Main Class  -------------------------- */
  
  export class LoadContextService {
    // Singleton instance
    private static instance: LoadContextService | null = null;
    
    private workbookManager: ClientWorkbookStateManager | null;
    private compressor: ClientSpreadsheetCompressor | null;
    
    // Query context builder for more efficient state capture
    private queryContextBuilder: QueryContextBuilder;
    // Chunk locator service for identifying relevant chunks
    private chunkLocator: ChunkLocatorService | null = null;
    // Embedding service for similarity search
    private embeddingService: EmbeddingService | null = null;
    // Whether advanced chunk location is enabled
    private useAdvancedChunkLocation: boolean = true;
    private anthropic: ClientAnthropicService;
    private mistral: MistralClientService;
    // Add to your class properties
    private metadataCacheByWorkbookId: Map<string, any> = new Map();


    /**
     * Private constructor to prevent direct instantiation
     */
    private constructor(params: {
      workbookStateManager?: ClientWorkbookStateManager | null;
      compressor?: ClientSpreadsheetCompressor | null;
      useAdvancedChunkLocation?: boolean;
      anthropic: ClientAnthropicService;
      mistral?: MistralClientService;
    }) {
      this.workbookManager = params.workbookStateManager ?? null;
      this.compressor = params.compressor ?? null;
      this.useAdvancedChunkLocation = params.useAdvancedChunkLocation ?? true;
      this.anthropic = params.anthropic;
      this.mistral = params.mistral || new MistralClientService(false);
      this.metadataCacheByWorkbookId = new Map();
  
      // Create the query context builder
      this.queryContextBuilder = new QueryContextBuilder(
        this.workbookManager, 
        this.workbookManager.getMetadataCache(),
        this.chunkLocator);
      
      // Initialize advanced chunk location components if enabled
      if (this.useAdvancedChunkLocation && this.workbookManager) {
        this.initializeChunkLocator();
      }
    }

    /**
     * Initialize the metadata cache for the current workbook
     * @param force Force a refresh of the cache even if it already exists
     */
    public setupCache(force: boolean = false): void {
        console.log(`%c LoadContextService: Setting up cache (force=${force})`, 'color: #8e44ad; font-weight: bold');
        this.queryContextBuilder.ensureAllSheetsCached(force);
    }
    
    /**
     * Get the singleton instance of LoadContextService
     * If it doesn't exist, create it with the provided parameters
     * @param params Parameters for creating a new instance if one doesn't exist
     * @returns The singleton LoadContextService instance
     */
    public static getInstance(params?: {
      workbookStateManager?: ClientWorkbookStateManager | null;
      compressor?: ClientSpreadsheetCompressor | null;
      useAdvancedChunkLocation?: boolean;
      anthropic: ClientAnthropicService;
      mistral?: MistralClientService;
    }): LoadContextService {
      if (!LoadContextService.instance) {
        if (!params) {
          throw new Error('LoadContextService not initialized. Must provide parameters for first initialization.');
        }
        console.log('%c Creating new LoadContextService singleton instance', 'color: #8e44ad; font-weight: bold');
        LoadContextService.instance = new LoadContextService(params);
      } else if (params) {
        console.log('%c Using existing LoadContextService singleton instance', 'color: #8e44ad; font-weight: bold');
      }
      
      return LoadContextService.instance;
    }
    
    /**
     * Reset the singleton instance (useful for testing or when switching workbooks)
     */
    public static resetInstance(): void {
      LoadContextService.instance = null;
      console.log('%c LoadContextService singleton instance reset', 'color: #8e44ad; font-weight: bold');
    }

    /**
   * Initialize the chunk locator service
   */
    private async initializeChunkLocator(): Promise<void> {
        console.log('%c Initializing advanced chunk location components', 'background: #8e44ad; color: #ecf0f1; font-size: 12px; padding: 2px 5px;');
        
        try {
        // Create and initialize the embedding service
        this.embeddingService = new EmbeddingService();
        await this.embeddingService.initialize();
        
        // Create the chunk locator service
        this.chunkLocator = new ChunkLocatorService({
            metadataCache: this.workbookManager.getMetadataCache(),
            embeddingStore: this.embeddingService,
            dependencyAnalyzer: this.workbookManager.getDependencyAnalyzer(),
            anthropicService: this.anthropic,
            mistralService: this.mistral,
            activeSheetName: this.workbookManager.getActiveSheetName()
        });
        
        // Attach the chunk locator to the query context builder
        this.queryContextBuilder.setChunkLocator(this.chunkLocator);
        
        console.log('%c Advanced chunk location components initialized successfully', 'color: #2ecc71');
        } catch (error) {
        console.error('Error initializing chunk locator:', error);
        console.log('%c Falling back to standard chunk identification', 'color: #e74c3c');
        this.useAdvancedChunkLocation = false;
        }
    }
  
    /**
     * Get the query context builder
     * @returns The query context builder instance
     */
    public getQueryContextBuilder(): QueryContextBuilder {
        return this.queryContextBuilder;
    }
    /**
     * Get the current workbook ID
     * @returns The workbook ID or a default value if not available
     */
    private getWorkbookId(): string {
        return this.workbookManager?.getWorkbookId() || 'default-workbook';
    }

    /**
     * Store metadata for the current workbook
     * @param metadata The metadata to store
     */
    public storeMetadataForCurrentWorkbook(metadata: any): void {
        const workbookId = this.getWorkbookId();
        this.metadataCacheByWorkbookId.set(workbookId, metadata);
        console.log(`Stored metadata for workbook: ${workbookId}`);
    }
  
    /**
     * Get metadata for the current workbook
     * @returns The metadata for the current workbook, or null if not found
     */
    public getMetadataForCurrentWorkbook(): any | null {
        const workbookId = this.getWorkbookId();
        const metadata = this.metadataCacheByWorkbookId.get(workbookId);
        return metadata || null;
    }
  
    /**
     * Check if metadata exists for the current workbook
     * @returns True if metadata exists for the current workbook
     */
    public hasMetadataForCurrentWorkbook(): boolean {
        const workbookId = this.getWorkbookId();
    return this.metadataCacheByWorkbookId.has(workbookId);
  }
}
