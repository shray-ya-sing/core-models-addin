/**
 * GreetingsService.ts
 * Service to detect and respond to simple greetings
 */

export class GreetingsService {
    // Array of simple greetings for quick response optimization
    private simpleGreetings: string[] = [
      'hi',
      'hi cori',
      'hi there',
      'hello',
      'hello cori',
      'hello there',
      'hey',
      'hey cori',
      'hey there',
      'good morning',
      'good afternoon',
      'good evening',
      'how are you',
      'how are you today',
      'how are you doing',
      'what are you up to',
      'what\'s up',
      'greetings',
      'what\'s going on',
      'what\'s the weather today',
      'what\'s your name',
      'what\'s the time',
      'where am i',
      'where are you',
      'good night',
      'how\'s it going',
      'how\'s it going today',
      'how\'s it hanging'
    ];
    
    // Standard response for simple greetings
    private standardGreetingResponse: string = "Hi there! Could I help you with any workbook related tasks today?";
  
    /**
    * Checks if a query is a simple greeting
     * @param query The user query to check
     * @returns True if the query is a simple greeting, false otherwise
     */
    public isGreeting(query: string): boolean {
      // Normalize the query (lowercase and trim)
      const normalizedQuery = query.toLowerCase().trim();
      
      // First check for exact matches
      if (this.simpleGreetings.some(greeting => normalizedQuery === greeting)) {
        return true;
      }
      
      // If no exact match, check for similarity using Levenshtein distance
      return this.simpleGreetings.some(greeting => {
        // Calculate Levenshtein distance
        const distance = this.levenshteinDistance(normalizedQuery, greeting);
        
        // Allow typos based on the length of the greeting
        // For short greetings (2-3 chars), allow 1 typo
        // For longer greetings, allow up to 2 typos
        const maxAllowedDistance = greeting.length <= 3 ? 1 : 2;
        
        return distance <= maxAllowedDistance;
      });
    }
    
    /**
     * Get the standard greeting response
     * @returns The standard greeting response
     */
    public getStandardResponse(): string {
      return this.standardGreetingResponse;
    }
    
    /**
     * Handle a greeting query and return the standard response
     * @param query The user query
     * @param streamingCallback Optional callback for streaming the response
     * @param processId Optional process ID for logging
     * @returns The standard greeting response
     */
    public handleGreeting(query: string, streamingCallback?: (text: string) => void, processId?: string): string {
      if (processId) {
        console.log(`%c Detected simple greeting (ID: ${processId.substring(0, 8)}) ${query}`, 'background: #27ae60; color: #fff; font-size: 12px; padding: 2px 5px;');
      } else {
        console.log('%c Detected simple greeting', 'background: #27ae60; color: #fff; font-size: 12px; padding: 2px 5px;');
      }
      
      // If streaming callback is provided, simulate streaming by sending the response
      // character by character with small delays
      if (streamingCallback) {
        // Send the first character immediately
        const response = this.standardGreetingResponse;
        let currentIndex = 0;
        
        // Function to stream the next character
        const streamNextChar = () => {
          if (currentIndex < response.length) {
            // Stream the next character
            streamingCallback(response.charAt(currentIndex));
            currentIndex++;
            
            // Schedule the next character with a small delay (15-30ms)
            setTimeout(streamNextChar, Math.random() * 15 + 15);
          }
        };
        
        // Start streaming
        streamNextChar();
      }
      
      return this.standardGreetingResponse;
    }
  
    /**
     * Calculates the Levenshtein distance between two strings
     * @param a First string
     * @param b Second string
     * @returns The Levenshtein distance
     */
    private levenshteinDistance(a: string, b: string): number {
      if (a.length === 0) return b.length;
      if (b.length === 0) return a.length;
    
      const matrix = [];
    
      // Initialize the matrix
      for (let i = 0; i <= b.length; i++) {
        matrix[i] = [i];
      }
      for (let j = 0; j <= a.length; j++) {
        matrix[0][j] = j;
      }
    
      // Fill the matrix
      for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
          const cost = a[j - 1] === b[i - 1] ? 0 : 1;
          matrix[i][j] = Math.min(
            matrix[i - 1][j] + 1,      // deletion
            matrix[i][j - 1] + 1,      // insertion
            matrix[i - 1][j - 1] + cost // substitution
          );
        }
      }
    
      return matrix[b.length][a.length];
    }
  }