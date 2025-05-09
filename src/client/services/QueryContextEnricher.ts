/**
 * Query Context Enricher
 * Enhances query context with chart metadata for improved LLM understanding
 */

import { QueryContext } from '../models/CommandModels';

/**
 * Service for enriching query context
 */
export class QueryContextEnricher {
  /**
   * Constructor
   */
  constructor() {
  }
  
  /**
   * Enrich a query context
   * @param context The original query context
   * @returns The enriched query context
   */
  public enrichQueryContext(context: QueryContext): QueryContext {
    // Clone the context to avoid modifying the original
    const enrichedContext = { ...context };

    return enrichedContext;
  }
}
