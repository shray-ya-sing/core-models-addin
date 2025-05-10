import { pipeline, env } from '@xenova/transformers';

// Set to use WASM backend for better browser compatibility
env.backends.onnx.wasm.numThreads = 1;
import axios from 'axios';

// Define the instruction metadata interface
interface InstructionMetadata {
  title: string;
  description: string;
  content: any;
}

// Cache for loaded models and instructions
let sentenceModel: any = null;
let modelLoading: Promise<any> | null = null;
let instructionsCache: InstructionMetadata[] = [];
let instructionsLoading: Promise<InstructionMetadata[]> | null = null;

let instructions = [{
    "title": "Analysis At Various Prices (AAVP) Table",
    "description": "Instructions for building an Analysis At Various Prices table for public companies in Excel financial models",
    "content": {
      "overview": "An AAVP table shows how key valuation metrics change across different stock price scenarios, helping analysts understand sensitivity to price changes",
      "prerequisites": [
        "Current stock price",
        "Shares outstanding",
        "Total debt",
        "Cash & equivalents",
        "Current EBITDA or earnings metric",
        "Current valuation multiples"
      ],
      "table_structure": {
        "columns": [
          "Price per share scenarios",
          "% Change from Current Price",
          "Implied Market Cap",
          "Implied Enterprise Value (EV)",
          "Implied EV/[Earnings Metric] Multiple",
          "Additional user-specified metrics (P/E, FCF yield, etc.)"
        ],
        "price_scenarios": {
          "default": "-20% to +20% in 5% increments",
          "custom_options": [
            "User-specified percentage range",
            "User-specified absolute values",
            "User-specified step size"
          ]
        }
      },
      "formulas": {
        "percent_change": "=([Price Scenario Cell] / [Current Price Cell]) - 1",
        "market_cap": "=[Price Scenario Cell] * [Shares Outstanding Cell]",
        "enterprise_value": "=[Implied Market Cap Cell] + [Total Debt Cell] - [Cash Cell]",
        "ev_multiple": "=[Implied Enterprise Value Cell] / [Earnings Metric Cell]",
        "pe_ratio": "=[Price Scenario Cell] / ([Net Income Cell] / [Shares Outstanding Cell])",
        "fcf_yield": "=([FCF Cell] / [Shares Outstanding Cell]) / [Price Scenario Cell]",
        "dividend_yield": "=[Annual Dividend Cell] / [Price Scenario Cell]"
      },
      "formatting": {
        "percent_change": "Percentage format",
        "monetary_values": "Currency format with appropriate precision",
        "multiples": "1-2 decimal places (or as specified)",
        "highlighting": "Conditional formatting for current price row",
        "styling": "Borders and shading as specified"
      },
      "integration": {
        "sensitivity_analysis": "Create named ranges for columns to reference in other model sections",
        "dynamic_updates": "Ensure formulas update if price scenarios change"
      },
      "keywords": [
        "AAVP",
        "sensitivity analysis",
        "valuation",
        "stock price",
        "multiple",
        "enterprise value",
        "market cap",
        "financial modeling",
        "equity analysis"
      ]
    }
  }]

/**
 * Calculate cosine similarity between two vectors
 */
function cosineSimilarity(vec1: number[], vec2: number[]) {
    const dotProduct = vec1.reduce((sum, value, index) => sum + value * vec2[index], 0);
    const magnitude1 = Math.sqrt(vec1.reduce((sum, value) => sum + value * value, 0));
    const magnitude2 = Math.sqrt(vec2.reduce((sum, value) => sum + value * value, 0));
    return dotProduct / (magnitude1 * magnitude2);
}

/**
 * Load all instruction JSON files
 */
async function loadInstructions(): Promise<InstructionMetadata[]> {
    // Create the instruction cache from the instructions array
    // convert instructions array to InstructionMetadata
    instructionsCache = instructions.map((instruction) => ({
        title: instruction.title,
        description: instruction.description,
        content: instruction.content
    }));
    return instructionsCache;
}

/**
 * Ensure the model is loaded
 */
async function ensureModelLoaded() {
    // Return existing model if available
    if (sentenceModel) {
        return sentenceModel;
    }
    
    // Return existing loading promise if one is in progress
    if (modelLoading) {
        return modelLoading;
    }
    
    // Start loading the model
    console.log('Loading sentence similarity model...');
    modelLoading = (async () => {
        try {
            sentenceModel = await pipeline('feature-extraction', 'xenova/all-MiniLM-L6-v2');
            console.log('Sentence similarity model loaded successfully');
            return sentenceModel;
        } catch (error) {
            console.error('Error loading sentence similarity model:', error);
            throw error;
        } finally {
            // Clear the loading promise
            modelLoading = null;
        }
    })();
    
    return modelLoading;
}

/**
 * Get similarity between two texts
 */
async function getSimilarity(text1: string, text2: string) {
    // Ensure model is loaded
    const model = await ensureModelLoaded();
    
    // Get embeddings for both texts
    // Feature extraction returns embeddings differently than sentence-similarity
    const embedding1 = await model(text1, { pooling: 'mean', normalize: true });
    const embedding2 = await model(text2, { pooling: 'mean', normalize: true });
    
    // Calculate cosine similarity
    const similarity = cosineSimilarity(embedding1.data, embedding2.data);
    
    return similarity;
}

/**
 * Result interface for instruction retrieval with similarity score
 */
export interface InstructionSearchResult {
    instruction: InstructionMetadata;
    similarity: number;
}

/**
 * Retrieve the most relevant instruction based on a query
 */
export async function retrieveRelevantInstructions(query: string): Promise<InstructionMetadata | null> {
    // Load instructions
    const instructions = await loadInstructions();
    if (instructions.length === 0) {
        return null;
    }
    
    // Find the most similar instruction
    let maxSimilarity = -1;
    let mostRelevant: InstructionMetadata | null = null;
    
    for (const instruction of instructions) {
        // Compare query with instruction title and description
        const combinedText = `${instruction.title}. ${instruction.description}`;
        const similarity = await getSimilarity(query, combinedText);
        
        if (similarity > maxSimilarity) {
            maxSimilarity = similarity;
            mostRelevant = instruction;
        }
    }
    
    return mostRelevant;
}

/**
 * Retrieve multiple relevant instructions based on a query, sorted by relevance
 * @param query The search query
 * @param limit Maximum number of results to return
 * @param threshold Minimum similarity score threshold (0-1)
 */
export async function retrieveMultipleRelevantInstructions(
    query: string,
    limit: number = 3,
    threshold: number = 0.5
): Promise<InstructionSearchResult[]> {
    // Load instructions
    const instructions = await loadInstructions();
    if (instructions.length === 0) {
        return [];
    }
    
    // Calculate similarity for all instructions
    const results: InstructionSearchResult[] = [];
    
    for (const instruction of instructions) {
        // Compare query with instruction title and description
        const combinedText = `${instruction.title}. ${instruction.description}`;
        const similarity = await getSimilarity(query, combinedText);
        
        if (similarity >= threshold) {
            results.push({
                instruction,
                similarity
            });
        }
    }
    
    // Sort by similarity (highest first) and limit results
    return results
        .sort((a, b) => b.similarity - a.similarity)
        .slice(0, limit);
}

/**
 * Test the instruction retrieval service with a sample query
 * This can be called from the console to verify the service is working
 */
export async function testInstructionRetrieval(query: string = 'How do I create a price analysis table?'): Promise<void> {
    console.log(`Testing instruction retrieval with query: "${query}"`);
    
    try {
        // Test loading instructions
        const instructions = await loadInstructions();
        console.log(`Loaded ${instructions.length} instructions:`);
        instructions.forEach(instruction => {
            console.log(`- ${instruction.title}`);
        });
        
        // Test retrieving relevant instructions
        console.log('\nRetrieving relevant instructions...');
        const results = await retrieveMultipleRelevantInstructions(query, 3, 0.3);
        
        if (results.length === 0) {
            console.log('No relevant instructions found.');
        } else {
            console.log(`Found ${results.length} relevant instructions:`);
            results.forEach((result, index) => {
                console.log(`\n${index + 1}. ${result.instruction.title} (Similarity: ${(result.similarity * 100).toFixed(1)}%)`);
                console.log(`   Description: ${result.instruction.description}`);
            });
        }
    } catch (error) {
        console.error('Error testing instruction retrieval:', error);
    }
}