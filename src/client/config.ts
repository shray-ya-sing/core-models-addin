/**
 * Client-side configuration
 * 
 * This file contains configuration values for the client-side application.
 * In a production environment, these would be loaded from environment variables
 * or a secure configuration service.
 */

export const config = {
  // API keys
  anthropicApiKey: process.env.ANTHROPIC_API_KEY || '',
  
  // API endpoints
  knowledgeBaseApiUrl: process.env.KNOWLEDGE_BASE_API_URL || 'http://localhost:8000/api/search/unified',
  
  // Model configuration
  anthropicModel: process.env.ANTHROPIC_MODEL || 'claude-3-opus-20240229',
  
  // Feature flags
  enableKnowledgeBase: process.env.ENABLE_KNOWLEDGE_BASE === 'true' || false,
  
  // Debug settings
  debugMode: process.env.DEBUG_MODE === 'true' || false,
};

// Log configuration in debug mode
if (config.debugMode) {
  console.log('Client configuration loaded:', {
    ...config,
    anthropicApiKey: config.anthropicApiKey ? '[REDACTED]' : 'Not set',
  });
}

export default config;
