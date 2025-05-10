import { pipeline } from '@xenova/transformers';
import { MetadataChunk } from '../../models/CommandModels';
import { EmbeddingStore, EmbeddingVector, SimilaritySearchResult } from './ChunkLocatorService';

// Configure the library to use the Transformers.js CDN for models
// Note: Configuration is done at runtime via the pipeline options

/**
 * Service for generating embeddings and performing semantic search on spreadsheet metadata
 * Implements the EmbeddingStore interface for integration with ChunkLocatorService
 */
export class EmbeddingService implements EmbeddingStore {
  private embeddingModel: any = null;
  private isInitialized: boolean = false;
  private embeddingCache: Map<string, EmbeddingVector> = new Map();
  private modelName: string = 'Xenova/all-MiniLM-L6-v2'; // Small, fast model (384 dimensions)
  
  /**
   * Initialize the embedding model
   */
  public async initialize(): Promise<void> {
    if (this.isInitialized) return;
    
    try {
      console.log(`🔄 [EmbeddingService] Initializing embedding model: ${this.modelName}`);
      this.embeddingModel = await pipeline('feature-extraction', this.modelName);
      this.isInitialized = true;
      console.log('✅ [EmbeddingService] Embedding model initialized successfully');
    } catch (error) {
      console.error('❌ [EmbeddingService] Failed to initialize embedding model:', error);
      throw error;
    }
  }
  
  /**
   * Convert a metadata chunk to a text representation for embedding
   */
  private chunkToText(chunk: MetadataChunk): string {
    // Extract data from the payload based on chunk type
    const payload = chunk.payload || {};
    const sheetName = payload.name || 'Unknown';
    const range = payload.range || 'full sheet';
    
    const parts = [
      `ID: ${chunk.id}`,
      `Sheet: ${sheetName}`,
      `Range: ${range}`,
      `Type: ${chunk.type}`,
    ];
    
    // Add summary if available
    if (chunk.summary) {
      parts.push(`Summary: ${chunk.summary}`);
    }
    
    // Add payload data if available
    if (payload) {
      // Extract values if available
      if (payload.values && Array.isArray(payload.values)) {
        const valueTexts = payload.values.flat()
          .filter(v => v !== null && v !== undefined)
          .map(v => String(v))
          .slice(0, 50); // Limit to avoid too long texts
        if (valueTexts.length > 0) {
          parts.push(`Values: ${valueTexts.join(', ')}`);
        }
      }
      
      // Extract formulas if available
      if (payload.formulas && Array.isArray(payload.formulas)) {
        const formulaTexts = payload.formulas.flat()
          .filter(f => f !== null && f !== undefined)
          .map(f => String(f))
          .slice(0, 20); // Limit to avoid too long texts
        if (formulaTexts.length > 0) {
          parts.push(`Formulas: ${formulaTexts.join(', ')}`);
        }
      }
    }
    
    return parts.join(' | ');
  }
  
  /**
   * Generate an embedding for a single text
   */
  public async generateEmbedding(text: string): Promise<EmbeddingVector> {
    if (!this.isInitialized) {
      await this.initialize();
    }
    
    // Check cache first
    const cacheKey = text.substring(0, 100); // Use first 100 chars as key
    if (this.embeddingCache.has(cacheKey)) {
      return this.embeddingCache.get(cacheKey);
    }
    
    // Generate embedding
    const output = await this.embeddingModel(text, { 
      pooling: 'mean', 
      normalize: true,
      // Configure the library to use the Transformers.js CDN for models
      use_cdn: true,
      revision: 'main'
    });
    
    // Convert to number array and cache
    const embedding = Array.from(output.data) as number[];
    this.embeddingCache.set(cacheKey, embedding);
    
    return embedding;
  }
  
  /**
   * Generate embeddings for multiple metadata chunks
   */
  public async generateChunkEmbeddings(chunks: MetadataChunk[]): Promise<Map<string, EmbeddingVector>> {
    if (!this.isInitialized) {
      await this.initialize();
    }
    
    const embeddingMap = new Map<string, EmbeddingVector>();
    
    // Process chunks in batches to avoid memory issues
    const batchSize = 10;
    for (let i = 0; i < chunks.length; i += batchSize) {
      const batch = chunks.slice(i, i + batchSize);
      const texts = batch.map(chunk => this.chunkToText(chunk));
      
      // Generate embeddings for the batch
      const outputs = await Promise.all(
        texts.map(text => this.generateEmbedding(text))
      );
      
      // Store in map using chunk id as key
      batch.forEach((chunk, index) => {
        embeddingMap.set(chunk.id, outputs[index]);
      });
    }
    
    return embeddingMap;
  }
  
  /**
   * Find chunks most similar to a query
   * Implements the EmbeddingStore interface method
   */
  public async findSimilarChunks(
    query: string, 
    chunks: MetadataChunk[], 
    topK: number = 5
  ): Promise<SimilaritySearchResult[]> {
    if (!this.isInitialized) {
      await this.initialize();
    }
    
    // Generate query embedding
    const queryEmbedding = await this.generateEmbedding(query);
    
    // Generate embeddings for chunks (or use cached ones)
    const chunkEmbeddings = await this.generateChunkEmbeddings(chunks);
    
    // Calculate similarity scores
    const results = chunks.map(chunk => {
      const chunkEmbedding = chunkEmbeddings.get(chunk.id);
      if (!chunkEmbedding) {
        return { chunkId: chunk.id, score: 0 };
      }
      
      // Calculate cosine similarity using the built-in methods
      const score = this.cosineSimilarity(queryEmbedding, chunkEmbedding);
      return { chunkId: chunk.id, score };
    });
    
    // Sort by score (descending) and take top K
    return results
      .sort((a, b) => b.score - a.score)
      .slice(0, topK);
  }
  
  /**
   * Calculate cosine similarity between two vectors
   */
  private cosineSimilarity(a: EmbeddingVector, b: EmbeddingVector): number {
    let dotProduct = 0;
    let normA = 0;
    let normB = 0;
    
    for (let i = 0; i < a.length; i++) {
      dotProduct += a[i] * b[i];
      normA += a[i] * a[i];
      normB += b[i] * b[i];
    }
    
    return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
  }
  
  /**
   * Clear the embedding cache
   * Implements the EmbeddingStore interface method
   */
  public clear(): void {
    this.embeddingCache.clear();
    console.log('🧹 [EmbeddingService] Embedding cache cleared');
  }
  
  /**
   * Get embedding for a specific chunk
   * Implements the EmbeddingStore interface method
   */
  public async getEmbedding(chunk: MetadataChunk, forceRefresh: boolean = false): Promise<EmbeddingVector> {
    const cacheKey = chunk.id;
    
    // Check cache first unless force refresh is requested
    if (!forceRefresh && this.embeddingCache.has(cacheKey)) {
      return this.embeddingCache.get(cacheKey);
    }
    
    // Generate text representation of the chunk
    const text = this.chunkToText(chunk);
    
    // Generate embedding
    const embedding = await this.generateEmbedding(text);
    
    // Cache the result
    this.embeddingCache.set(cacheKey, embedding);
    
    return embedding;
  }
}
