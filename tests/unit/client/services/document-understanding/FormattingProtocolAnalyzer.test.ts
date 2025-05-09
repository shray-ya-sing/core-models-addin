// Mock the Anthropic SDK first before any imports
jest.mock('@anthropic-ai/sdk', () => {
  return {
    __esModule: true,
    default: jest.fn().mockImplementation(() => ({
      messages: {
        create: jest.fn().mockResolvedValue({
          content: [
            {
              type: 'text',
              text: '```json\n{"colorCoding":{"calculations":"#000000","hardcodedValues":"#0000FF","inputs":"#E6E6E6"},"numberFormatting":{"currency":"$#,##0.00","percentage":"0.00%"},"borderStyles":{"tables":"thin","totals":"double"},"chartFormatting":{"chartTypes":{"preferred":["columnClustered","line"]}},"generalObservations":["Consistent color coding"],"modelQualityAssessment":{"consistencyScore":0.9,"clarityScore":0.8,"professionalismScore":0.85,"comments":["Well structured"]},"confidenceScore":0.9}\n```'
            }
          ]
        })
      }
    }))
  };
});

// Now import after mocks are setup
import { FormattingProtocolAnalyzer, FormattingProtocol, WorkbookFormattingMetadata, ThemeColors } from '../../../../../src/client/services/document-understanding/FormattingProtocolAnalyzer';
import { ClientAnthropicService } from '../../../../../src/client/services/ClientAnthropicService';

// Mock ClientAnthropicService to avoid actual SDK initialization
jest.mock('../../../../../src/client/services/ClientAnthropicService', () => {
  // Create mock response data
  const mockResponseJson = '{"colorCoding":{"calculations":"#000000","hardcodedValues":"#0000FF","inputs":"#E6E6E6"},"numberFormatting":{"currency":"$#,##0.00","percentage":"0.00%"},"borderStyles":{"tables":"thin","totals":"double"},"chartFormatting":{"chartTypes":{"preferred":["columnClustered","line"]}},"generalObservations":["Consistent color coding"],"modelQualityAssessment":{"consistencyScore":0.9,"clarityScore":0.8,"professionalismScore":0.85,"comments":["Well structured"]},"confidenceScore":0.9}';
  
  // Create mock client
  const mockAnthropicClient = {
    messages: {
      create: jest.fn().mockResolvedValue({
        content: [
          {
            type: 'text',
            text: '```json\n' + mockResponseJson + '\n```'
          }
        ]
      })
    }
  };
  
  return {
    ClientAnthropicService: jest.fn().mockImplementation(() => ({
      getClient: jest.fn().mockReturnValue(mockAnthropicClient),
      getModel: jest.fn().mockReturnValue('claude-3-opus-20240229'),
      extractJsonFromMarkdown: jest.fn((text) => {
        const match = text.match(/```json\s*([\s\S]*?)\s*```/);
        return match ? match[1] : null;
      })
    })),
    ModelType: {
      Advanced: 'advanced'
    }
  };
});

// Skip Excel.run tests for now
jest.mock('../../../../mocks/office.js.mock', () => ({}), { virtual: true });

describe('FormattingProtocolAnalyzer', () => {
  let analyzer: FormattingProtocolAnalyzer;
  let anthropicService: ClientAnthropicService;

  beforeEach(() => {
    // Create the anthropic service with a mock API key
    anthropicService = new ClientAnthropicService('test-api-key');
    analyzer = new FormattingProtocolAnalyzer(anthropicService);
  });

  describe('analyzeFormattingProtocol', () => {
    it('should analyze formatting protocol using LLM', async () => {
      // Mock data
      const images = ['base64image1', 'base64image2'];
      const formattingMetadata: WorkbookFormattingMetadata = {
        themeColors: {
          background1: '#FFFFFF',
          background2: '#F2F2F2',
          text1: '#000000',
          text2: '#666666',
          accent1: '#4472C4',
          accent2: '#ED7D31',
          accent3: '#A5A5A5',
          accent4: '#FFC000',
          accent5: '#5B9BD5',
          accent6: '#70AD47',
          hyperlink: '#0563C1',
          followedHyperlink: '#954F72'
        },
        sheets: []
      };

      // Call the method
      const result = await analyzer.analyzeFormattingProtocol(images, formattingMetadata);

      // Verify the result
      expect(result).toBeDefined();
      expect(result.colorCoding).toBeDefined();
      expect(result.colorCoding.calculations).toBe('#000000');
      expect(result.colorCoding.hardcodedValues).toBe('#0000FF');
      expect(result.colorCoding.inputs).toBe('#E6E6E6');
      expect(result.numberFormatting).toBeDefined();
      expect(result.numberFormatting.currency).toBe('$#,##0.00');
      expect(result.numberFormatting.percentage).toBe('0.00%');
      expect(result.borderStyles).toBeDefined();
      expect(result.borderStyles.tables).toBe('thin');
      expect(result.borderStyles.totals).toBe('double');
      expect(result.chartFormatting).toBeDefined();
      expect(result.chartFormatting.chartTypes?.preferred).toContain('columnClustered');
      expect(result.generalObservations).toContain('Consistent color coding');
      expect(result.modelQualityAssessment).toBeDefined();
      expect(result.modelQualityAssessment.consistencyScore).toBe(0.9);
      expect(result.confidenceScore).toBe(0.9);

      // Verify the API call
      expect(anthropicService.getClient().messages.create).toHaveBeenCalledWith(
        expect.objectContaining({
          model: 'claude-3-opus-20240229',
          messages: expect.arrayContaining([
            expect.objectContaining({
              role: 'user',
              content: expect.arrayContaining([
                expect.objectContaining({
                  type: 'text',
                  text: expect.stringContaining('Please analyze')
                }),
                expect.objectContaining({
                  type: 'image',
                  source: expect.objectContaining({
                    data: 'base64image1'
                  })
                }),
                expect.objectContaining({
                  type: 'image',
                  source: expect.objectContaining({
                    data: 'base64image2'
                  })
                }),
                expect.objectContaining({
                  type: 'text',
                  text: expect.stringContaining('Formatting Metadata')
                })
              ])
            })
          ])
        })
      );

      // Verify JSON extraction was called
      expect(anthropicService.extractJsonFromMarkdown).toHaveBeenCalled();
    });

    it('should throw an error if JSON extraction fails', async () => {
      // Mock data
      const images = ['base64image1'];
      const formattingMetadata: WorkbookFormattingMetadata = {
        themeColors: {} as ThemeColors,
        sheets: []
      };

      // Mock extractJsonFromMarkdown to return null
      (anthropicService.extractJsonFromMarkdown as jest.Mock).mockReturnValueOnce(null);

      // Expect the method to throw
      await expect(analyzer.analyzeFormattingProtocol(images, formattingMetadata))
        .rejects.toThrow('No valid JSON formatting protocol found in LLM response');
    });

    it('should throw an error if JSON parsing fails', async () => {
      // Mock data
      const images = ['base64image1'];
      const formattingMetadata: WorkbookFormattingMetadata = {
        themeColors: {} as ThemeColors,
        sheets: []
      };

      // Mock extractJsonFromMarkdown to return invalid JSON
      (anthropicService.extractJsonFromMarkdown as jest.Mock).mockReturnValueOnce('{invalid:json}');

      // Expect the method to throw
      await expect(analyzer.analyzeFormattingProtocol(images, formattingMetadata))
        .rejects.toThrow('Failed to parse formatting protocol from LLM response');
    });
  });
});
