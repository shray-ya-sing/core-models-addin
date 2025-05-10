import { MetadataChunk } from '../../models/CommandModels';

/**
 * Vector embedding type
 */
export type EmbeddingVector = number[];

/**
 * Metadata for a stored embedding
 */
export interface EmbeddingMetadata {
  // ID of the chunk
  chunkId: string;
  // ETag of the chunk when embedding was created
  etag: string;
  // Timestamp when embedding was created
  createdAt: number;
  // Type of embedding model used
  modelType: string;
}

/**
 * Stored embedding for a chunk
 */
export interface StoredEmbedding {
  // Embedding vector
  vector: EmbeddingVector;
  // Metadata about the embedding
  metadata: EmbeddingMetadata;
}

/**
 * Configuration for the EmbeddingStore
 */
export interface EmbeddingStoreConfig {
  // Whether to use a local or remote embedding model
  useLocalModel: boolean;
  // Type of embedding model to use
  modelType: 'openai' | 'local';
  // Whether to persist embeddings to local storage
  persistToLocalStorage: boolean;
  // Max size of local storage for embeddings (in bytes)
  maxLocalStorageSize: number;
  // Local storage key prefix
  localStorageKeyPrefix: string;
}

/**
 * Default configuration for the EmbeddingStore
 */
const DEFAULT_CONFIG: EmbeddingStoreConfig = {
  useLocalModel: true, // Default to local model for initial phase
  modelType: 'local',
  persistToLocalStorage: true,
  maxLocalStorageSize: 5 * 1024 * 1024, // 5MB
  localStorageKeyPrefix: 'excel_addin_embedding_',
};

/**
 * Similarity search result
 */
export interface SimilaritySearchResult {
  chunkId: string;
  score: number;
}

/**
 * Service for computing and storing vector embeddings
 */
export class EmbeddingStore {
  private embeddings: Map<string, StoredEmbedding> = new Map();
  private config: EmbeddingStoreConfig;
  private isInitialized: boolean = false;

  constructor(config?: Partial<EmbeddingStoreConfig>) {
    this.config = { ...DEFAULT_CONFIG, ...(config || {}) };
  }

  /**
   * Initialize the embedding store
   * Loads persisted embeddings from local storage if available
   */
  public async initialize(): Promise<void> {
    if (this.isInitialized) return;

    console.log('%c Initializing EmbeddingStore', 'color: #3498db');

    // Load persisted embeddings if enabled
    if (this.config.persistToLocalStorage) {
      await this.loadFromLocalStorage();
    }

    this.isInitialized = true;
  }

  /**
   * Get an embedding for a chunk
   * @param chunk The chunk to get or compute embedding for
   * @param forceRefresh Whether to force recomputation of the embedding
   * @returns The embedding vector
   */
  public async getEmbedding(chunk: MetadataChunk, forceRefresh: boolean = false): Promise<EmbeddingVector> {
    // Ensure we're initialized
    if (!this.isInitialized) {
      await this.initialize();
    }

    const chunkId = chunk.id;
    const etag = chunk.etag;

    // Check if we have a valid cached embedding
    if (!forceRefresh && this.embeddings.has(chunkId)) {
      const storedEmbedding = this.embeddings.get(chunkId)!;
      
      // Check if the chunk has changed since the embedding was computed
      if (storedEmbedding.metadata.etag === etag) {
        return storedEmbedding.vector;
      }
    }

    // Need to compute a new embedding
    const vector = await this.computeEmbedding(chunk);
    
    // Store the new embedding
    const newEmbedding: StoredEmbedding = {
      vector,
      metadata: {
        chunkId,
        etag,
        createdAt: Date.now(),
        modelType: this.config.modelType,
      },
    };
    
    this.embeddings.set(chunkId, newEmbedding);
    
    // Persist to local storage if enabled
    if (this.config.persistToLocalStorage) {
      this.saveToLocalStorage(chunkId, newEmbedding);
    }
    
    return vector;
  }

  /**
   * Compute an embedding for a chunk
   * @param chunk The chunk to compute an embedding for
   * @returns The embedding vector
   */
  private async computeEmbedding(chunk: MetadataChunk): Promise<EmbeddingVector> {
    // Create a text representation of the chunk
    const chunkText = this.getTextRepresentation(chunk);
    
    if (this.config.useLocalModel) {
      // For now, simulate a local embedding with a simple hash-based vector
      return this.computeSimpleEmbedding(chunkText);
    } else {
      // In the future, implement call to OpenAI embeddings API
      console.log('%c Remote embedding API not yet implemented', 'color: #e74c3c');
      return this.computeSimpleEmbedding(chunkText);
    }
  }

  /**
   * Get a text representation of a chunk for embedding
   * @param chunk The chunk to represent as text
   * @returns Text representation of the chunk
   */
  private getTextRepresentation(chunk: MetadataChunk): string {
    if (chunk.type === 'sheet' && chunk.payload) {
      const sheet = chunk.payload;
      let text = `Sheet: ${sheet.name}\n`;
      
      // Add summary if available
      if (sheet.summary) {
        text += `Summary: ${sheet.summary}\n`;
      }
      
      // Add anchors
      if (sheet.anchors && Array.isArray(sheet.anchors)) {
        text += 'Key cells:\n';
        for (const anchor of sheet.anchors.slice(0, 10)) { // Limit to first 10 anchors
          if (anchor.value && anchor.address) {
            text += `- ${anchor.address}: ${anchor.value}\n`;
          }
        }
      }
      
      // Add some sample values
      if (sheet.values && Array.isArray(sheet.values)) {
        text += 'Sample values:\n';
        for (const value of sheet.values.slice(0, 5)) { // Limit to first 5 values
          if (value.value && value.address) {
            text += `- ${value.address}: ${value.value}\n`;
          }
        }
      }
      
      return text;
    } else if (chunk.type === 'range' && chunk.payload) {
      const range = chunk.payload;
      let text = `Range: ${range.sheet}!${range.address}\n`;
      
      // Add description if available
      if (range.description) {
        text += `Description: ${range.description}\n`;
      }
      
      // Add values if available
      if (range.values && Array.isArray(range.values)) {
        text += 'Values:\n';
        for (const row of range.values.slice(0, 3)) { // Limit to first 3 rows
          text += `- ${row.join(', ')}\n`;
        }
      }
      
      return text;
    } else {
      // Default representation
      return `Chunk ${chunk.id} of type ${chunk.type}`;
    }
  }

  /**
   * Compute a simple embedding vector based on text
   * This is a placeholder until a proper embedding model is implemented
   * @param text The text to embed
   * @returns A simple embedding vector
   */
  private computeSimpleEmbedding(text: string): EmbeddingVector {
    // Simple placeholder embedding (128-dim vector based on character frequencies)
    // This is not a proper embedding, just a placeholder
    const vector: number[] = new Array(128).fill(0);
    
    // Normalize the text
    const normalizedText = text.toLowerCase();
    
    // Fill the vector with character frequencies
    for (let i = 0; i < normalizedText.length; i++) {
      const char = normalizedText.charCodeAt(i);
      vector[char % 128] += 1;
    }
    
    // Normalize the vector to unit length
    const magnitude = Math.sqrt(vector.reduce((sum, val) => sum + val * val, 0));
    if (magnitude > 0) {
      for (let i = 0; i < vector.length; i++) {
        vector[i] /= magnitude;
      }
    }
    
    return vector;
  }

  /**
   * Find the most similar chunks to a query
   * @param query The query text
   * @param chunks The chunks to search
   * @param topK The number of most similar chunks to return
   * @returns Array of similarity search results
   */
  public async findSimilarChunks(
    query: string,
    chunks: MetadataChunk[],
    topK: number = 5
  ): Promise<SimilaritySearchResult[]> {
    // Ensure we're initialized
    if (!this.isInitialized) {
      await this.initialize();
    }

    // Compute embedding for the query
    const queryVector = await this.computeSimpleEmbedding(query);
    
    // Get embeddings for all chunks
    const results: SimilaritySearchResult[] = [];
    for (const chunk of chunks) {
      try {
        const chunkVector = await this.getEmbedding(chunk);
        const score = this.computeCosineSimilarity(queryVector, chunkVector);
        
        results.push({
          chunkId: chunk.id,
          score,
        });
      } catch (error) {
        console.error(`Error computing similarity for chunk ${chunk.id}:`, error);
      }
    }
    
    // Sort by score (descending) and take top K
    return results
      .sort((a, b) => b.score - a.score)
      .slice(0, topK);
  }

  /**
   * Compute cosine similarity between two vectors
   * @param a First vector
   * @param b Second vector
   * @returns Cosine similarity score (0-1)
   */
  private computeCosineSimilarity(a: number[], b: number[]): number {
    if (a.length !== b.length) {
      throw new Error(`Vector dimensions do not match: ${a.length} vs ${b.length}`);
    }
    
    let dotProduct = 0;
    let aMagnitude = 0;
    let bMagnitude = 0;
    
    for (let i = 0; i < a.length; i++) {
      dotProduct += a[i] * b[i];
      aMagnitude += a[i] * a[i];
      bMagnitude += b[i] * b[i];
    }
    
    aMagnitude = Math.sqrt(aMagnitude);
    bMagnitude = Math.sqrt(bMagnitude);
    
    if (aMagnitude === 0 || bMagnitude === 0) {
      return 0;
    }
    
    return dotProduct / (aMagnitude * bMagnitude);
  }

  /**
   * Load embeddings from local storage
   */
  private async loadFromLocalStorage(): Promise<void> {
    try {
      // Get all keys that match our prefix
      const keys = Object.keys(localStorage).filter(key => 
        key.startsWith(this.config.localStorageKeyPrefix)
      );
      
      console.log(`%c Loading ${keys.length} embeddings from local storage`, 'color: #3498db');
      
      // Load each embedding
      let loadedCount = 0;
      for (const key of keys) {
        try {
          const storedValue = localStorage.getItem(key);
          if (storedValue) {
            const embedding = JSON.parse(storedValue) as StoredEmbedding;
            const chunkId = key.replace(this.config.localStorageKeyPrefix, '');
            this.embeddings.set(chunkId, embedding);
            loadedCount++;
          }
        } catch (err) {
          console.error(`Error loading embedding from key ${key}:`, err);
        }
      }
      
      console.log(`%c Successfully loaded ${loadedCount} embeddings from local storage`, 'color: #2ecc71');
    } catch (error) {
      console.error('Error loading embeddings from local storage:', error);
    }
  }

  /**
   * Save an embedding to local storage
   * @param chunkId The chunk ID
   * @param embedding The embedding to save
   */
  private saveToLocalStorage(chunkId: string, embedding: StoredEmbedding): void {
    try {
      const key = `${this.config.localStorageKeyPrefix}${chunkId}`;
      const value = JSON.stringify(embedding);
      
      localStorage.setItem(key, value);
    } catch (error) {
      console.error(`Error saving embedding for chunk ${chunkId} to local storage:`, error);
    }
  }

  /**
   * Clear all embeddings from memory and local storage
   */
  public clear(): void {
    this.embeddings.clear();
    
    if (this.config.persistToLocalStorage) {
      try {
        // Remove all keys that match our prefix
        const keys = Object.keys(localStorage).filter(key => 
          key.startsWith(this.config.localStorageKeyPrefix)
        );
        
        for (const key of keys) {
          localStorage.removeItem(key);
        }
        
        console.log(`%c Cleared ${keys.length} embeddings from local storage`, 'color: #e74c3c');
      } catch (error) {
        console.error('Error clearing embeddings from local storage:', error);
      }
    }
  }
}
