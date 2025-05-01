export enum KnowledgeBaseStatus {
  Idle = 'idle',
  Searching = 'searching',
  Success = 'success',
  Error = 'error'
}

export interface KnowledgeBaseEvent {
  status: KnowledgeBaseStatus;
  message: string;
  data?: any;
  error?: Error;
}

export type KnowledgeBaseEventListener = (event: KnowledgeBaseEvent) => void;

/**
 * Client-side service for interacting with the knowledge base API
 */
export class ClientKnowledgeBaseService {
  private apiUrl: string;
  private eventListeners: KnowledgeBaseEventListener[] = [];
  private currentStatus: KnowledgeBaseStatus = KnowledgeBaseStatus.Idle;

  constructor(apiUrl: string = 'http://localhost:8000/api/search/unified') {
    this.apiUrl = apiUrl;
  }

  /**
   * Set the API URL for the knowledge base
   * @param url The API URL
   */
  public setApiUrl(url: string): void {
    this.apiUrl = url;
  }
  
  /**
   * Add an event listener for knowledge base status changes
   * @param listener The event listener to add
   * @returns A function to remove the listener
   */
  public addEventListener(listener: KnowledgeBaseEventListener): () => void {
    this.eventListeners.push(listener);
    return () => {
      this.eventListeners = this.eventListeners.filter(l => l !== listener);
    };
  }
  
  /**
   * Get the current status of the knowledge base
   * @returns The current status
   */
  public getStatus(): KnowledgeBaseStatus {
    return this.currentStatus;
  }
  
  /**
   * Emit an event to all listeners
   * @param event The event to emit
   */
  private emitEvent(event: KnowledgeBaseEvent): void {
    this.currentStatus = event.status;
    this.eventListeners.forEach(listener => listener(event));
  }

  /**
   * Search the knowledge base for relevant documents
   * @param query The search query
   * @param limit Maximum number of results to return
   * @returns Search results or null if error
   */
  public async search(query: string, limit: number = 5): Promise<any> {
    // Emit searching event
    this.emitEvent({
      status: KnowledgeBaseStatus.Searching,
      message: 'Searching knowledge base...'
    });
    
    try {
      // Add a short timeout to make the searching status visible
      const requestStartTime = Date.now();
      
      const response = await fetch(`${this.apiUrl}?query=${encodeURIComponent(query)}&limit=${limit}`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        },
        // Add timeout to prevent long hangs
        signal: AbortSignal.timeout(10000) // 10 second timeout
      });

      if (!response.ok) {
        const errorText = await response.text().catch(() => 'No error details available');
        throw new Error(`Knowledge base API error: ${response.status} - ${errorText}`);
      }

      // If response was too fast, add a small delay for better UX
      const elapsedTime = Date.now() - requestStartTime;
      if (elapsedTime < 500) {
        await new Promise(resolve => setTimeout(resolve, 500 - elapsedTime));
      }

      // Parse the response
      const data = await response.json();
      
      // Check for valid data structure
      if (!data || !Array.isArray(data.results)) {
        throw new Error('Invalid response format from knowledge base');
      }
      
      // Emit success event
      this.emitEvent({
        status: KnowledgeBaseStatus.Success,
        message: `Found ${data.results.length} results in knowledge base`,
        data
      });
      
      return data;
    } catch (error) {
      console.error('Error searching knowledge base:', error);
      
      // Emit error event
      this.emitEvent({
        status: KnowledgeBaseStatus.Error,
        message: `Failed to search knowledge base: ${error.message || 'Unknown error'}`,
        error
      });
      
      // Return empty results on error
      return {
        results: [],
        total_results: 0,
        tables_found: 0,
        content_types_found: []
      };
    }
  }

  /**
   * Extract relevant data points from the knowledge base
   * @param query The context query
   * @param limit Maximum number of results to return
   * @returns Extracted data points
   */
  public async extractDataPoints(query: string, limit: number = 5): Promise<any> {
    try {
      const searchResults = await this.search(query, limit);
      
      // Extract numerical data from the search results
      const dataPoints = this.extractNumericalData(searchResults.results);
      
      return {
        dataPoints,
        searchResults
      };
    } catch (error) {
      console.error('Error extracting data points:', error);
      return {
        dataPoints: [],
        searchResults: {
          results: [],
          total_results: 0
        }
      };
    }
  }

  /**
   * Extract numerical data from search results
   * @param results Search results
   * @returns Array of numerical data points
   */
  private extractNumericalData(results: any[]): any[] {
    const dataPoints: any[] = [];
    
    for (const result of results) {
      if (!result.content) continue;
      
      // Look for patterns like "$X million", "X%", "X dollars"
      const content = result.content;
      const currencyMatches = content.matchAll(/\$(\d+(?:\.\d+)?)\s*(million|billion|thousand)?/g);
      const percentMatches = content.matchAll(/(\d+(?:\.\d+)?)\s*%/g);
      const numberMatches = content.matchAll(/(\d+(?:\.\d+)?)\s*(dollars|euros|pounds)/g);
      
      // Process currency matches
      for (const match of currencyMatches) {
        let value = parseFloat(match[1]);
        if (match[2]) {
          switch (match[2].toLowerCase()) {
            case 'million': value *= 1000000; break;
            case 'billion': value *= 1000000000; break;
            case 'thousand': value *= 1000; break;
          }
        }
        
        // Get some context around the match
        const startIndex = Math.max(0, content.indexOf(match[0]) - 50);
        const endIndex = Math.min(content.length, content.indexOf(match[0]) + match[0].length + 50);
        const context = content.substring(startIndex, endIndex);
        
        dataPoints.push({
          type: 'currency',
          value,
          originalText: match[0],
          context,
          source: result.metadata?.title || 'Unknown'
        });
      }
      
      // Process percentage matches
      for (const match of percentMatches) {
        const value = parseFloat(match[1]);
        
        // Get some context around the match
        const startIndex = Math.max(0, content.indexOf(match[0]) - 50);
        const endIndex = Math.min(content.length, content.indexOf(match[0]) + match[0].length + 50);
        const context = content.substring(startIndex, endIndex);
        
        dataPoints.push({
          type: 'percentage',
          value,
          originalText: match[0],
          context,
          source: result.metadata?.title || 'Unknown'
        });
      }
      
      // Process number matches
      for (const match of numberMatches) {
        const value = parseFloat(match[1]);
        
        // Get some context around the match
        const startIndex = Math.max(0, content.indexOf(match[0]) - 50);
        const endIndex = Math.min(content.length, content.indexOf(match[0]) + match[0].length + 50);
        const context = content.substring(startIndex, endIndex);
        
        dataPoints.push({
          type: 'number',
          value,
          originalText: match[0],
          context,
          source: result.metadata?.title || 'Unknown'
        });
      }
    }
    
    return dataPoints;
  }
}
