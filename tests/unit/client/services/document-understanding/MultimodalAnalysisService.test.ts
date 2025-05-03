// Mock the Anthropic SDK first before any imports
jest.mock('@anthropic-ai/sdk', () => {
  return {
    __esModule: true,
    default: jest.fn().mockImplementation(() => ({
      messages: {
        create: jest.fn().mockResolvedValue({
          content: [{ type: 'text', text: 'Mock response from Anthropic' }]
        })
      }
    }))
  };
});

// Set up global Excel object that is used directly in the code
// This needs to be done before importing the service
global.Excel = {
  run: jest.fn((callback) => Promise.resolve(callback({
    workbook: {
      worksheets: {
        getCount: jest.fn().mockResolvedValue(3),
        load: jest.fn(),
        items: [{ name: 'Sheet1' }]
      }
    },
    sync: jest.fn().mockResolvedValue(undefined)
  })))
};

// Now import after mocks are setup
import { MultimodalAnalysisService } from '../../../../../src/client/services/document-understanding/MultimodalAnalysisService';
import { ClientAnthropicService } from '../../../../../src/client/services/ClientAnthropicService';

// Mock the WorkbookUtils module
jest.mock('../../../../../src/client/services/document-understanding/WorkbookUtils', () => ({
  performMultimodalAnalysis: jest.fn().mockResolvedValue(['base64image1', 'base64image2'])
}));

// Mock ClientAnthropicService to avoid actual SDK initialization
jest.mock('../../../../../src/client/services/ClientAnthropicService', () => {
  return {
    ClientAnthropicService: jest.fn().mockImplementation(() => ({
      getClient: jest.fn(),
      getModel: jest.fn().mockReturnValue('claude-3-7-sonnet-20250219'),
      extractJsonFromMarkdown: jest.fn((text) => '{"test":"value"}'),
      classifyQueryType: jest.fn().mockResolvedValue({
        query_type: 'workbook_analysis',
        steps: []
      })
    })),
    ModelType: {
      Advanced: 'advanced'
    }
  };
});

describe('MultimodalAnalysisService', () => {
  let service: MultimodalAnalysisService;
  
  beforeEach(() => {
    // Reset mocks
    jest.clearAllMocks();
    
    // Create the service with mock dependencies
    // Pass a dummy API key to satisfy the constructor parameter requirement
    const anthropicService = new ClientAnthropicService('test-api-key');
    service = new MultimodalAnalysisService('http://test-endpoint.com', anthropicService);
  });

  describe('analyzeActiveWorkbook', () => {
    it('should perform multimodal analysis and return images', async () => {
      // Import the mocked function to verify it's called
      const { performMultimodalAnalysis } = require('../../../../../src/client/services/document-understanding/WorkbookUtils');
      
      // Call the method with no options
      const result = await service.analyzeActiveWorkbook();
      
      // Verify the performMultimodalAnalysis was called with the right arguments
      expect(performMultimodalAnalysis).toHaveBeenCalledWith('http://test-endpoint.com', undefined);
      
      // Verify the result structure
      expect(result).toBeDefined();
      expect(result.images).toEqual(['base64image1', 'base64image2']);
      expect(result.metadata).toBeDefined();
      // We can't verify the timestamp as it's dynamic, but we can check for its presence
      expect(result.metadata.timestamp).toBeDefined();
    });
    
    it('should handle analysis options correctly', async () => {
      // Import the mocked function to verify it's called
      const { performMultimodalAnalysis } = require('../../../../../src/client/services/document-understanding/WorkbookUtils');
      
      // Call with some options
      const options = {
        sheets: ['Sheet1', 'Sheet2']
      };
      
      await service.analyzeActiveWorkbook(options);
      
      // Verify options were passed correctly
      expect(performMultimodalAnalysis).toHaveBeenCalledWith('http://test-endpoint.com', options);
    });
  });
});
