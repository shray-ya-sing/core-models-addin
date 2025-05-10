import { MetadataChunk } from '../../models/CommandModels';
import { ClientAnthropicService } from '../llm/ClientAnthropicService';
import { WorkbookMetadataCache } from './WorkbookMetadataCache';
import { RangeDependencyAnalyzer } from './RangeDependencyAnalyzer';
import { ChatHistoryMessage } from '../request-processing/ClientQueryProcessor';
import { MistralClientService } from '../llm/MistralClientService';
// Forward declarations for EmbeddingStore which will be implemented later
export type EmbeddingVector = number[];

export interface EmbeddingStore {
  initialize(): Promise<void>;
  getEmbedding(chunk: MetadataChunk, forceRefresh?: boolean): Promise<EmbeddingVector>;
  findSimilarChunks(query: string, chunks: MetadataChunk[], topK?: number): Promise<SimilaritySearchResult[]>;
  clear(): void;
}

export interface SimilaritySearchResult {
  chunkId: string;
  score: number;
}

/**
 * Result of chunk location process
 */
export interface ChunkLocatorResult {
  // Array of chunk IDs that are relevant to the query
  chunkIds: string[];
  // Detailed information about what was found
  details: {
    sheets: string[];     // Sheet names
    ranges: string[];     // Range references (e.g. "Sheet1!A1:B10")
    charts: string[];     // Chart titles
    semantics: Record<string, string>; // Semantic mappings (e.g. "Revenue" -> "Sheet1!B5:B10")
  };
  // Confidence scores for each chunk ID
  confidenceScores: Map<string, number>;
  // Whether the location process used LLM assistance
  usedLLM: boolean;
}

/**
 * Configuration for the ChunkLocatorService
 */
export interface ChunkLocatorConfig {
  // Whether to use LLM for assistance in locating chunks
  enableLLM: boolean;
  // Number of candidate chunks to retrieve using embeddings
  maxCandidates: number;
  // Minimum confidence score for a chunk to be included
  confidenceThreshold: number;
  // Whether to use naive LLM-based selection instead of embeddings
  useNaiveLLMSelection: boolean;
}

/**
 * Default configuration for the ChunkLocatorService
 */
const DEFAULT_CONFIG: ChunkLocatorConfig = {
  enableLLM: true, // Start with LLM disabled for initial phase
  maxCandidates: 20,
  confidenceThreshold: 0.5,
  useNaiveLLMSelection: true, // Enable the naive LLM selection approach as requested
};

/**
 * Service for locating relevant chunks based on a query
 * Uses a multi-stage approach:
 * 1. Rule-based matching for explicit mentions
 * 2. Embedding-based similarity search
 * 3. LLM-based ranking and selection (optional)
 * 4. Dependency expansion
 */
export class ChunkLocatorService {
  private metadataCache: WorkbookMetadataCache;
  private embeddingStore: EmbeddingStore;
  private dependencyAnalyzer: RangeDependencyAnalyzer;
  private anthropicService: ClientAnthropicService;
  private config: ChunkLocatorConfig;
  private activeSheetName: string | null = null;
  private chatHistory: ChatHistoryMessage[] = [];
  private mistralService: MistralClientService;

  constructor(params: {
    metadataCache: WorkbookMetadataCache;
    embeddingStore: EmbeddingStore;
    dependencyAnalyzer: RangeDependencyAnalyzer;
    anthropicService?: ClientAnthropicService;
    mistralService?: MistralClientService;
    config?: Partial<ChunkLocatorConfig>;
    activeSheetName?: string;
    chatHistory?: ChatHistoryMessage[];
  }) {
    this.metadataCache = params.metadataCache;
    this.embeddingStore = params.embeddingStore;
    this.dependencyAnalyzer = params.dependencyAnalyzer;
    this.anthropicService = params.anthropicService;
    this.mistralService = params.mistralService;
    this.config = { ...DEFAULT_CONFIG, ...(params.config || {}) };
    this.activeSheetName = params.activeSheetName || null;
    this.chatHistory = params.chatHistory || [];
  }

  /**
   * Set the active sheet name
   * @param activeSheet The active sheet name
   */
  public setActiveSheet(activeSheet: string | null): void {
    this.activeSheetName = activeSheet;
  }

  /**
   * Locate chunks relevant to a query
   * @param query The query text
   * @returns Promise with the chunk locator result
   */
  public async locateChunks(query: string, chatHistory: Array<{role: string, content: string}>): Promise<ChunkLocatorResult> {
    // Enhanced debug logging to track query through the method chain
    console.log(
      `%c ChunkLocatorService: Locating chunks for query: "${query}"`,
      'background: #8e44ad; color: #ecf0f1; font-size: 12px; padding: 2px 5px;'
    );
    
    // Log query length and first/last characters to check for whitespace issues
    console.log(`%c Query length: ${query.length}, First char code: ${query.charCodeAt(0)}, Last char code: ${query.charCodeAt(query.length-1)}`, 'color: #3498db;');

    // Initialize the result
    const result: ChunkLocatorResult = {
      chunkIds: [],
      details: {
        sheets: [],
        ranges: [],
        charts: [],
        semantics: {},
      },
      confidenceScores: new Map<string, number>(),
      usedLLM: false,
    };
    
    // Always try LLM selection first when available
    if (this.config.useNaiveLLMSelection && this.anthropicService) {
      console.log('%c Using LLM-based sheet selection', 'color: #8e44ad');
      await this.performNaiveLLMSelection(query, result, chatHistory);
      
      // If we got results from the LLM, use them; otherwise fall back to rule-based matching
      if (result.chunkIds.length > 0) {
        console.log('%c LLM successfully identified relevant sheets', 'color: #2ecc71');
        this.expandDependencies(result);
        return result;
      }
      console.log('%c Naive LLM selection produced no results, falling back', 'color: #e74c3c');
    }

    else{
      console.log('%c Using rule-based matching', 'color: #8e44ad');
      // 1. Rule-based matching (fast explicit matching)
      const explicitMatches = await this.performExplicitMatching(query, result);
      
      // If we have high-confidence explicit matches, we can skip the more expensive matching
      if (explicitMatches.highConfidenceMatch) {
        console.log('%c Using high-confidence explicit matches only', 'color: #27ae60');
        this.expandDependencies(result);
        return result;
      }

      // 2. Embedding-based similarity search
      await this.performEmbeddingSearch(query, result);

      // 3. LLM-based ranking and selection (optional)
      if (this.config.enableLLM && this.anthropicService) {
        await this.performLLMRanking(query, result);
      }

      // 4. Dependency expansion
      this.expandDependencies(result);
    }

    // Log the final results
    console.log(
      `%c ChunkLocatorService: Located ${result.chunkIds.length} chunks`,
      'background: #27ae60; color: #ecf0f1; font-size: 12px; padding: 2px 5px;'
    );
    
    if (result.details.sheets.length > 0) {
      console.log(`%c Located sheets: ${result.details.sheets.join(', ')}`, 'color: #3498db');
    }
    
    if (result.details.ranges.length > 0) {
      console.log(`%c Located ranges: ${result.details.ranges.join(', ')}`, 'color: #3498db');
    }
    
    if (result.details.charts.length > 0) {
      console.log(`%c Located charts: ${result.details.charts.join(', ')}`, 'color: #3498db');
    }

    return result;
  }

  /**
   * Perform explicit matching for quickly identifying relevant chunks
   * @param query The query text
   * @param result The result object to populate
   * @returns Object indicating if high-confidence matches were found
   */
  private async performExplicitMatching(
    query: string,
    result: ChunkLocatorResult
  ): Promise<{highConfidenceMatch: boolean}> {
    console.log('%c Performing explicit matching...', 'color: #3498db');
  
    // Normalize the query - clean up special characters for better matching
    const normalizedQuery = query.toLowerCase().replace(/[\.\,\/#!$%\^&\*;:{}=\-_`~()]/g, ' ');
  
    // Remove pattern matching and let LLM handle everything
  
    // Track if we found high confidence matches
    let highConfidenceMatch = false;
  
    // Get all sheet chunks
    const sheetChunks = this.metadataCache.getAllSheetChunks();
    const activeSheetId = this.activeSheetName ? `Sheet:${this.activeSheetName}` : null;
    let hasHighConfidenceMatch = false;
  
    // Create a version of the query with reference terms removed
    const queryWithoutReferenceTerms = normalizedQuery
      .replace(/\s+tab\b/g, '')
      .replace(/\s+sheet\b/g, '')
      .replace(/\s+worksheet\b/g, '');
    
    // 1. Check for explicit sheet name mentions
    for (const chunk of sheetChunks) {
      if (!chunk.payload || !chunk.payload.name) continue;
      
      const sheetName = chunk.payload.name;
      const sheetNameLower = sheetName.toLowerCase();
      
      // 1.1 Exact sheet name match (highest relevance)
      const exactRegex = new RegExp(`\\b${sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
      if (exactRegex.test(query)) {
        console.log(`%c Query explicitly mentions sheet: ${sheetName}`, 'color: #3498db');
        result.chunkIds.push(chunk.id);
        result.details.sheets.push(sheetName);
        result.confidenceScores.set(chunk.id, 1.0); // High confidence
        hasHighConfidenceMatch = true;
        continue;
      }
      
      // 1.2 Check for sheet name with tab/sheet terminology
      const sheetWithRefTerms = new RegExp(`\\b${sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s+(tab|sheet|worksheet)\\b`, 'i');
      if (sheetWithRefTerms.test(query)) {
        console.log(`%c Query explicitly mentions sheet (with term): ${sheetName}`, 'color: #3498db');
        result.chunkIds.push(chunk.id);
        result.details.sheets.push(sheetName);
        result.confidenceScores.set(chunk.id, 1.0); // High confidence
        hasHighConfidenceMatch = true;
        continue;
      }
      
      // 1.3 Partial sheet name match
      if (normalizedQuery.includes(sheetNameLower) || queryWithoutReferenceTerms.includes(sheetNameLower)) {
        console.log(`%c Query contains partial match for sheet: ${sheetName}`, 'color: #3498db');
        result.chunkIds.push(chunk.id);
        result.details.sheets.push(sheetName);
        result.confidenceScores.set(chunk.id, 0.8); // Good confidence
        hasHighConfidenceMatch = true;
        continue;
      }
    }

    // 2. Check for cell range references
    const rangeRegex = /\b([A-Z]+[0-9]+:[A-Z]+[0-9]+|[A-Z]+[0-9]+)\b/g;
    const rangeMatches = query.match(rangeRegex);
    
    if (rangeMatches && rangeMatches.length > 0) {
      console.log(`%c Query contains cell range references: ${rangeMatches.join(', ')}`, 'color: #3498db');
      
      // If we have specific ranges but no chunks identified yet, add the active sheet
      if (result.chunkIds.length === 0 && activeSheetId) {
        console.log(`%c Using active sheet for range references: ${this.activeSheetName}`, 'color: #3498db');
        result.chunkIds.push(activeSheetId);
        result.details.sheets.push(this.activeSheetName!);
        result.confidenceScores.set(activeSheetId, 0.7); // Fairly high confidence
      }
      
      // Add the range references to the result
      for (const range of rangeMatches) {
        // If no sheet specified in the range, assume active sheet
        if (!range.includes('!') && this.activeSheetName) {
          const fullRange = `${this.activeSheetName}!${range}`;
          result.details.ranges.push(fullRange);
        } else {
          result.details.ranges.push(range);
        }
      }
    }

    // 3. Check for potential sheet reference patterns
    if (result.chunkIds.length === 0) {
      // Look for potential sheet name patterns like: "X tab", "X sheet", "sheet X", etc.
      const sheetPatterns = [
        /\b(.+?)\s+tab\b/i,       // "Financial Model tab"
        /\b(.+?)\s+sheet\b/i,      // "Financial Model sheet"
        /\bsheet\s+(.+?)\b/i,      // "sheet Financial Model"
        /\btab\s+(.+?)\b/i,        // "tab Financial Model"
        /\bworksheet\s+(.+?)\b/i,  // "worksheet Financial Model"
        /\bin\s+(.+?)\b/i          // "in Financial Model"
      ];
      
      // Check each pattern
      for (const pattern of sheetPatterns) {
        const match = query.match(pattern);
        if (match && match[1]) {
          const potentialSheetName = match[1].trim();
          console.log(`%c Found potential sheet reference: "${potentialSheetName}"`, 'color: #f39c12');
          
          // Check if this potential sheet name matches any sheet
          for (const chunk of sheetChunks) {
            if (!chunk.payload || !chunk.payload.name) continue;
            
            const sheetName = chunk.payload.name;
            // Check if the potential sheet name is similar to this sheet
            if (sheetName.toLowerCase().includes(potentialSheetName.toLowerCase()) || 
                potentialSheetName.toLowerCase().includes(sheetName.toLowerCase())) {
              console.log(`%c Matched potential sheet reference to actual sheet: ${sheetName}`, 'color: #2ecc71');
              result.chunkIds.push(chunk.id);
              result.details.sheets.push(sheetName);
              result.confidenceScores.set(chunk.id, 0.75); // Good confidence
              hasHighConfidenceMatch = true;
              break; // Found a match
            }
          }
          
          // If we found a match, no need to check other patterns
          if (hasHighConfidenceMatch) break;
        }
      }
    }

    // 4. If still no chunks identified, use active sheet
    if (result.chunkIds.length === 0 && activeSheetId) {
      console.log(`%c No relevant sheets identified, using active sheet: ${this.activeSheetName}`, 'color: #e74c3c');
      result.chunkIds.push(activeSheetId);
      result.details.sheets.push(this.activeSheetName!);
      result.confidenceScores.set(activeSheetId, 0.5); // Moderate confidence
    }

    return { highConfidenceMatch: hasHighConfidenceMatch };
  }

  /**
   * Perform embedding-based similarity search
   * @param query The query text
   * @param result The result object to populate
   */
  private async performEmbeddingSearch(
    query: string,
    result: ChunkLocatorResult
  ): Promise<void> {
    // This will be implemented in the next phase when we add the embedding store
    console.log('%c Embedding search not yet implemented', 'color: #e74c3c');
    
    // We'll implement this in the next phase
    // For now, just log that we would have searched for the query
    console.log(`%c Would search for query: "${query}"`, 'color: #f39c12');
    console.log('%c Would update result with embedding search results', 'color: #f39c12');
    
    // Placeholder to track that we went through this path
    // This also fixes the TS6133 error (result is declared but never read)
    result.usedLLM = false; // Indicate that we didn't use LLM for this result
  }

  /**
   * Perform LLM-based ranking and selection
   * @param query The query text
   * @param result The result object to populate
   */
  private async performLLMRanking(
    query: string,
    result: ChunkLocatorResult
  ): Promise<void> {
    // This will be implemented in a later phase
    console.log('%c LLM ranking not yet implemented', 'color: #e74c3c');
    
    // We'll implement this in a later phase
    // For now, just log that we would have used the LLM to rank results
    console.log(`%c Would use LLM to rank chunks for query: "${query}"`, 'color: #f39c12');
    console.log('%c Would update result with LLM-ranked chunks', 'color: #f39c12');
    
    // Mark that we used the LLM in the result
    result.usedLLM = true;
  }
  
  /**
   * Perform naive LLM-based sheet selection using Anthropic
   * @param query The user's query
   * @param result The result object to populate
   */
  private async performNaiveLLMSelection(
    query: string,
    result: ChunkLocatorResult,
    chatHistory: Array<{role: string, content: string}>
  ): Promise<void> {
    if (!this.anthropicService) {
      console.log('%c No Anthropic service available for LLM selection', 'color: #e74c3c');
      return;
    }
    
    // Debug logging to verify the query value
    console.log(`%c LLM Sheet Selection - Query: "${query}"`, 'background: #8e44ad; color: white; font-weight: bold;');
    console.log(`%c LLM Sheet Selection - Chat History Length: ${chatHistory.length}`, 'color: #3498db;');
    
    try {
      // Get all available sheets from the metadata cache
      const allSheets = this.metadataCache.getAllSheetChunks();
      
      // Create a list of sheets with summaries for the LLM
      const availableSheets = allSheets.map(chunk => ({
        name: chunk.payload?.name || chunk.id.replace('Sheet:', ''),
        summary: chunk.payload?.summary || ''
      }));
      
      console.log(`%c Found ${availableSheets.length} sheets to analyze with LLM`, 'color: #3498db');
      
      // Use the Anthropic service to select relevant sheets
      //uncomment this when anthropic is ready
      //const selectedSheetNames = await this.anthropicService.selectRelevantSheets(
      //  query,
      //  availableSheets,
      //  chatHistory
      //);

      // Use mistral to select relevant sheets
      const selectedSheetNames = await this.mistralService.selectRelevantSheets(
        query,
        availableSheets,
        chatHistory
      );
      
      if (selectedSheetNames.length === 0) {
        console.log('%c LLM did not select any sheets', 'color: #e74c3c');
        return;
      }
      
      console.log(
        `%c LLM selected ${selectedSheetNames.length} sheets: ${selectedSheetNames.join(', ')}`,
        'color: #2ecc71'
      );
      
      // Map sheet names back to chunk IDs and add to result
      for (const sheetName of selectedSheetNames) {
        // Find the matching chunk ID
        const matchingChunk = allSheets.find(chunk => 
          chunk.payload?.name === sheetName || 
          chunk.id === `Sheet:${sheetName}`
        );
        
        if (matchingChunk) {
          result.chunkIds.push(matchingChunk.id);
          result.details.sheets.push(sheetName);
          result.confidenceScores.set(matchingChunk.id, 0.9); // High confidence from LLM
        }
      }
      
      // Mark that we used LLM in the result
      result.usedLLM = true;
    } catch (error) {
      console.error('Error in naive LLM sheet selection:', error);
    }
  }
  
  /**
   * Expand dependencies for the identified chunks
   * @param result The result object to update with dependencies
   */
  private expandDependencies(result: ChunkLocatorResult): void {
    if (result.chunkIds.length === 0) return;
    
    console.log('%c Expanding dependencies for identified chunks', 'color: #3498db');
    
    // Add direct dependencies
    const directDependencies = new Set<string>();
    for (const chunkId of result.chunkIds) {
      const deps = this.dependencyAnalyzer.getDependencyChunks(chunkId);
      deps.forEach(depId => directDependencies.add(depId));
    }
    
    // Add transitive dependencies
    const allDependencies = this.dependencyAnalyzer.getTransitiveDependencies(result.chunkIds);
    
    // Log dependency information
    if (directDependencies.size > 0) {
      console.log(`%c Direct dependencies: ${Array.from(directDependencies).join(', ')}`, 'color: #3498db');
    }
    
    if (allDependencies.size > directDependencies.size) {
      console.log(
        `%c Also including transitive dependencies: ${
          Array.from(allDependencies)
            .filter(id => !directDependencies.has(id))
            .join(', ')
        }`, 
        'color: #3498db'
      );
    }
    
    // Add dependencies to result if not already included
    for (const depId of allDependencies) {
      if (!result.chunkIds.includes(depId)) {
        result.chunkIds.push(depId);
        
        // Set confidence scores for dependencies
        const isDirectDep = directDependencies.has(depId);
        result.confidenceScores.set(depId, isDirectDep ? 0.6 : 0.4);
        
        // Add to details if it's a sheet
        if (depId.startsWith('Sheet:')) {
          const sheetName = depId.replace('Sheet:', '');
          if (!result.details.sheets.includes(sheetName)) {
            result.details.sheets.push(sheetName);
          }
        }
      }
    }
  }
}
