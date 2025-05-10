/**
 * Formatting Protocol Analyzer for Excel workbooks
 * Facade that coordinates formatting metadata extraction and LLM analysis
 */
import { ClientAnthropicService, ModelType } from '../llm/ClientAnthropicService';
import { FormattingMetadataExtractor } from './FormattingMetadataExtractor';
import { FormattingProtocol, WorkbookFormattingMetadata } from './FormattingModels';
import { ExcelImageService } from './ExcelImageService';

/**
 * Regular expression to validate base64 strings
 * This regex checks for a valid base64 character set and proper padding
 */
const BASE64_REGEX = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;

/**
 * PNG file signature in base64 (first 8 bytes of a PNG file encoded as base64)
 * This corresponds to the PNG magic number: 89 50 4E 47 0D 0A 1A 0A
 */
const PNG_SIGNATURE_BASE64_PREFIX = 'iVBORw';

/**
 * System prompt for the formatting protocol analysis
 */
const FORMATTING_PROTOCOL_SYSTEM_PROMPT = `You are an expert Excel financial modeling analyst specializing in institutional-grade financial model formatting. Your task is to analyze Excel workbook images and formatting metadata to identify patterns and conventions used in the financial model.

You must identify the following aspects of the formatting protocol:

1. Color coding conventions
   - Identify colors used for different cell types (calculations, hardcoded values, inputs, etc.)
   - Note any color patterns for positive/negative values
   - Identify header/title formatting

2. Number formatting conventions
   - Currency formats
   - Percentage formats
   - Date formats
   - Negative number formatting

3. Border styling
   - Table borders
   - Section dividers
   - Total/subtotal row formatting

4. Chart formatting
   - Preferred chart types for different data
   - Title, legend, and axis formatting
   - Data label conventions

5. Font usage
   - Header fonts and styles
   - Body text fonts and styles
   - Title fonts and styles

6. Table formatting
   - Header row styling
   - Data row styling (alternating colors?)
   - Total row styling

7. Financial statement formatting
   - Income statement conventions
   - Balance sheet conventions
   - Cash flow statement conventions
   - Supporting schedule conventions

8. Workbook structure
   - Tab ordering/grouping
   - Section dividers within sheets
   - Input vs. calculation vs. output areas

9. Cover page formatting
   - Title styling
   - Company information placement
   - Date formatting

10. Scenario / Sensitivity table formatting
  - Usually have highlights or outlining to demarcate different scenarios
  - Usually laid out as a table with columns and rows for different variables under different scenarios. 
  - Usually, the middle column and middle row represent the "main" case while the others represent "sensitized" cases
  - Usually, the middle column and middle row are bolded and have a distinct background color

You MUST respond with a valid JSON object that follows this exact schema. DO NOT include any explanatory text, markdown formatting, or code block syntax. The response must be a valid JSON object that can be parsed directly.

Schema:

{
  "colorCoding": {
    "calculations": "#HEXCODE",
    "hardcodedValues": "#HEXCODE",
    "linkedValues": "#HEXCODE",
    "inputs": "#HEXCODE",
    "headers": "#HEXCODE",
    "totals": "#HEXCODE",
    "subtotals": "#HEXCODE",
    "errors": "#HEXCODE",
    "negativeValues": "#HEXCODE",
    "custom": {}
  },
  "numberFormatting": {
    "currency": "format pattern",
    "percentage": "format pattern",
    "date": "format pattern",
    "general": "format pattern",
    "negativeNumbers": "format pattern",
    "custom": {}
  },
  "borderStyles": {
    "tables": "description",
    "sections": "description",
    "totals": "description",
    "subtotals": "description",
    "schedules": "description",
    "custom": {}
  },
  "chartFormatting": {
    "chartTypes": {
      "preferred": ["chart type 1", "chart type 2"],
      "financial": ["chart type for financial data"],
      "custom": []
    },
    "title": {
      "hasTitle": true,
      "text": "typical title format",
      "position": "position description",
      "format": {
        "font": {
          "name": "font name",
          "size": 12,
          "bold": true,
          "color": "#HEXCODE"
        },
        "alignment": "alignment description"
      }
    },
    "legend": {
      "position": "position description",
      "format": {
        "font": {
          "name": "font name",
          "size": 10,
          "bold": false,
          "color": "#HEXCODE"
        }
      }
    },
    "dataLabels": {
      "visible": true,
      "position": "position description",
      "format": {
        "font": {
          "name": "font name",
          "size": 9,
          "bold": false,
          "color": "#HEXCODE"
        },
        "numberFormat": "format pattern"
      }
    },
    "axes": {
      "categoryAxis": {
        "title": {
          "visible": true,
          "text": "typical title format",
          "format": {
            "font": {
              "name": "font name",
              "size": 10,
              "bold": true,
              "color": "#HEXCODE"
            }
          }
        },
        "gridlines": false,
        "format": {
          "font": {
            "name": "font name",
            "size": 9,
            "bold": false,
            "color": "#HEXCODE"
          }
        }
      },
      "valueAxis": {
        "title": {
          "visible": true,
          "text": "typical title format",
          "format": {
            "font": {
              "name": "font name",
              "size": 10,
              "bold": true,
              "color": "#HEXCODE"
            }
          }
        },
        "gridlines": true,
        "format": {
          "font": {
            "name": "font name",
            "size": 9,
            "bold": false,
            "color": "#HEXCODE"
          },
          "numberFormat": "format pattern"
        }
      }
    },
    "seriesFormatting": {
      "colors": ["#HEXCODE1", "#HEXCODE2"],
      "markers": {
        "visible": true,
        "style": "style description"
      },
      "lineStyles": ["style description"]
    }
  },
  "fontUsage": {
    "headers": {
      "name": "font name",
      "size": 12,
      "bold": true,
      "color": "#HEXCODE"
    },
    "body": {
      "name": "font name",
      "size": 11,
      "bold": false,
      "color": "#HEXCODE"
    },
    "titles": {
      "name": "font name",
      "size": 14,
      "bold": true,
      "color": "#HEXCODE"
    },
    "custom": {}
  },
  "tableFormatting": {
    "headerRow": {
      "fillColor": "#HEXCODE",
      "fontColor": "#HEXCODE",
      "fontBold": true,
      "borders": "description"
    },
    "dataRows": {
      "alternatingColors": true,
      "evenRowColor": "#HEXCODE",
      "oddRowColor": "#HEXCODE"
    },
    "totalRow": {
      "fillColor": "#HEXCODE",
      "fontColor": "#HEXCODE",
      "fontBold": true,
      "borders": "description"
    }
  },
  "scheduleFormatting": {
    "incomeStatement": "description",
    "balanceSheet": "description",
    "cashFlow": "description",
    "debtSchedule": "description",
    "capex": "description",
    "depreciation": "description",
    "taxSchedule": "description",
    "workingCapital": "description",
    "custom": {}
  },
  "workbookStructure": {
    "tabOrdering": "description",
    "tabGrouping": "description",
    "sectionDividers": "description",
    "inputSections": "description",
    "outputSections": "description",
    "calculationSections": "description"
  },
  "scenarioFormatting": {
    "sensitivityTables": {
      "layout": "description",
      "highlighting": "description",
      "baseCase": {
        "position": "description",
        "formatting": "description"
      }
    },
    "scenarioManager": {
      "used": true,
      "structure": "description"
    },
    "dataTables": {
      "used": true,
      "structure": "description"
    }
  },
  "coverPageFormatting": {
    "title": {
      "font": {
        "name": "font name",
        "size": 16,
        "bold": true,
        "color": "#HEXCODE"
      },
      "alignment": "alignment description"
    },
    "subtitle": {
      "font": {
        "name": "font name",
        "size": 14,
        "bold": true,
        "color": "#HEXCODE"
      },
      "alignment": "alignment description"
    },
    "logo": {
      "position": "position description",
      "size": "size description"
    },
    "companyInfo": {
      "font": {
        "name": "font name",
        "size": 12,
        "bold": false,
        "color": "#HEXCODE"
      },
      "alignment": "alignment description"
    },
    "date": {
      "format": "format pattern",
      "position": "position description"
    }
  },
  "workbookType": {
    "financialModel": true,
    "threeStatementModel": false,
    "dcfModel": false,
    "lboModel": false,
    "mergersModel": false,
    "budgetModel": false,
    "forecastModel": false,
    "operationalModel": false,
    "dashboardModel": false,
    "custom": "description if applicable"
  },
  "yearSuffixes": {
    "actual": "A",
    "projected": "P",
    "budget": "B"
  }
}

Be specific and detailed in your analysis. Reference exact hex codes and formatting patterns from the metadata. Compare the observed formatting to institutional best practices for financial models. If you cannot determine a particular aspect with confidence, indicate this in your analysis.

ANTI-PATTERNS TO AVOID:
1. DO NOT include any explanatory text outside the JSON object
2. DO NOT wrap the JSON in markdown code blocks
3. DO NOT include any comments within the JSON
4. DO NOT use placeholders like "#HEXCODE" - use actual values from the metadata
5. DO NOT include properties that aren't in the schema
6. DO NOT omit required properties from the schema

Your response must be a single, valid, parseable JSON object and nothing else.`;

/**
 * Analyzer for Excel workbook formatting protocols
 */
export class FormattingProtocolAnalyzer {
  private anthropicService: ClientAnthropicService;
  private metadataExtractor: FormattingMetadataExtractor;
  private excelImageService: ExcelImageService;
  
  /**
   * Constructor
   * @param anthropicService The Anthropic service for LLM analysis
   * @param excelImageService Optional Excel image service for capturing workbook images
   */
  constructor(
    anthropicService: ClientAnthropicService,
    excelImageService?: ExcelImageService
  ) {
    this.anthropicService = anthropicService;
    this.metadataExtractor = new FormattingMetadataExtractor();
    this.excelImageService = excelImageService || new ExcelImageService();
  }
  
  /**
   * Validates if a string is a valid base64 encoded PNG image
   * @param base64String The base64 string to validate
   * @returns True if the string is a valid base64 encoded PNG image, false otherwise
   */
  private isValidBase64PngImage(base64String: string): boolean {
    try {
      // Check if the string is empty or null
      if (!base64String) {
        console.warn('Base64 string is empty or null');
        return false;
      }
      
      // Remove data URL prefix if present
      let cleanBase64 = base64String;
      if (base64String.startsWith('data:image/png;base64,')) {
        cleanBase64 = base64String.substring('data:image/png;base64,'.length);
      }
      
      // Check if the string matches the base64 pattern
      if (!BASE64_REGEX.test(cleanBase64)) {
        console.warn('String does not match base64 pattern');
        return false;
      }
      
      // Check if the string starts with the PNG signature
      if (!cleanBase64.startsWith(PNG_SIGNATURE_BASE64_PREFIX)) {
        console.warn('String does not start with PNG signature');
        return false;
      }
      
      // Additional validation: check if the decoded length is reasonable
      // PNG files should be at least a few hundred bytes
      const decodedLength = Math.floor(cleanBase64.length * 0.75);
      if (decodedLength < 100) {
        console.warn('Decoded base64 length is too small for a valid PNG');
        return false;
      }
      
      return true;
    } catch (error) {
      console.error('Error validating base64 PNG image:', error);
      return false;
    }
  }
  
  /**
   * Extracts formatting metadata from the active workbook
   * @returns Promise with workbook formatting metadata
   */
  public async extractFormattingMetadata(): Promise<WorkbookFormattingMetadata> {
    return this.metadataExtractor.extractFormattingMetadata();
  }
  
  /**
   * Prepares a message with workbook images and formatting metadata for LLM analysis
   * @returns A message object ready for the LLM API
   */
  public async prepareFormattingProtocolMessage(): Promise<any> {
    try {
      // Get workbook images
      const images = await this.excelImageService.captureWorkbookImages();
      
      // Extract formatting metadata
      const formattingMetadata = await this.extractFormattingMetadata();
      
      // Prepare the message content with images and metadata
      const message = {
        type: 'text',
        text: 'Analyze the formatting protocol of this Excel workbook.'
      };
      
      // Add formatting metadata
      const metadataMessage = {
        type: 'text',
        text: `Formatting Metadata:\n\`\`\`json\n${JSON.stringify(formattingMetadata, null, 2)}\n\`\`\`\n\nPlease analyze ${images.length > 0 ? 'the images and' : 'the'} metadata to identify the formatting protocol used in this workbook. Follow the schema provided in the system prompt.`
      };
      
      return {
        message,
        images,
        metadataMessage,
        formattingMetadata
      };
    } catch (error) {
      console.error('Error preparing formatting protocol message:', error);
      throw error;
    }
  }
  
  /**
   * Analyzes the workbook formatting protocol using LLM
   * @returns A FormattingProtocol object containing the analysis results
   */
  public async analyzeFormattingProtocol(): Promise<FormattingProtocol> {
    try {
      // Prepare message with images and formatting metadata
      const messageData = await this.prepareFormattingProtocolMessage();
      
      // Check if we have any images in the message
      const hasImages = messageData.messages.some(msg => msg.type === 'image');
      if (!hasImages) {
        console.warn('No images available for formatting protocol analysis. Proceeding with metadata only.');
      }
      
      // Try with Anthropic first (up to 2 retries)
      let formattingProtocol: FormattingProtocol | null = null;
      let error: Error | null = null;
      
      // Try Anthropic up to 3 times (initial + 2 retries)
      for (let attempt = 0; attempt < 3; attempt++) {
        try {
          // If not the first attempt, add a feedback message to the system prompt
          let updatedSystemPrompt = FORMATTING_PROTOCOL_SYSTEM_PROMPT;
          if (attempt > 0) {
            updatedSystemPrompt = `IMPORTANT: Previous response was invalid. Please ensure you return ONLY valid JSON without any markdown formatting, explanatory text, or code block syntax. The response must be a valid JSON object that can be parsed directly.\n\n${FORMATTING_PROTOCOL_SYSTEM_PROMPT}`;
          }
          
          // Call the Anthropic API to analyze the formatting protocol
          const response = await this.callAnthropicApi(
            messageData,
            updatedSystemPrompt
          );
          
          // Extract text from the response
          const responseText = this.extractTextFromResponse(response);
          
          // Use the ClientAnthropicService's extractJsonFromMarkdown method to extract JSON
          const extractedJson = this.anthropicService.extractJsonFromMarkdown(responseText);
          
          if (extractedJson) {
            try {
              formattingProtocol = JSON.parse(extractedJson);
              // Successfully parsed, break out of retry loop
              break;
            } catch (e) {
              console.error(`Attempt ${attempt + 1}: Error parsing JSON from Anthropic response:`, e);
              error = new Error('Failed to parse formatting protocol from Anthropic response');
            }
          } else {
            console.error(`Attempt ${attempt + 1}: No valid JSON formatting protocol found in Anthropic response`);
            error = new Error('No valid JSON formatting protocol found in Anthropic response');
          }
        } catch (e) {
          console.error(`Attempt ${attempt + 1}: Error calling Anthropic API:`, e);
          error = e instanceof Error ? e : new Error(String(e));
        }
      }
      
      // If Anthropic attempts failed, try OpenAI as fallback
      if (!formattingProtocol) {
        try {
          console.log('Anthropic attempts failed, falling back to OpenAI...');
          
          // TODO: Replace with actual OpenAI API call
          // This is a placeholder for the OpenAI API call with response_format: 'json_object'
          // const openAiResponse = await this.openAiService.createChatCompletion({
          //   model: 'gpt-4-turbo',
          //   messages: [
          //     { role: 'system', content: FORMATTING_PROTOCOL_SYSTEM_PROMPT },
          //     { role: 'user', content: JSON.stringify(messageData.formattingMetadata) }
          //   ],
          //   response_format: { type: 'json_object' }
          // });
          // 
          // const openAiResponseText = openAiResponse.choices[0].message.content;
          // formattingProtocol = JSON.parse(openAiResponseText);
          
          // For now, just throw the last error since OpenAI integration is not implemented
          throw error || new Error('Failed to analyze formatting protocol with Anthropic');
        } catch (e) {
          console.error('Error with OpenAI fallback:', e);
          throw e instanceof Error ? e : new Error(String(e));
        }
      }
      
      return formattingProtocol;
    } catch (error) {
      console.error('Error analyzing formatting protocol:', error);
      throw error;
    }
  }
  
  /**
   * Calls the Anthropic API to analyze the formatting protocol
   * @param messageData The message data with images and metadata
   * @param systemPrompt The system prompt to use
   * @returns The response from the Anthropic API
   */
  private async callAnthropicApi(messageData: any, systemPrompt: string): Promise<any> {
    // Create message content array
    const messageContent: any[] = [
      messageData.message
    ];
    
    // Add images to the message content, but only if they're valid
    let validImageCount = 0;
    for (const imageBase64 of messageData.images) {
      // Check if the image is valid before adding it
      if (this.isValidBase64PngImage(imageBase64)) {
        messageContent.push({
          type: 'image',
          source: {
            type: 'base64',
            media_type: 'image/png',
            data: imageBase64
          }
        });
        validImageCount++;
      } else {
        console.warn('Skipping invalid base64 PNG image in FormattingProtocolAnalyzer');
      }
    }
    
    // Log how many valid images were added
    console.log(`Added ${validImageCount} valid images to the message content out of ${messageData.images.length} total images`);
    
    // If no valid images were found, add a message indicating this
    if (validImageCount === 0 && messageData.images.length > 0) {
      messageContent.push({
        type: 'text',
        text: "Note: No valid worksheet images could be processed. The analysis will be based solely on the formatting metadata."
      });
    }
    
    // Add metadata message
    messageContent.push(messageData.metadataMessage);
    
    // Get the Anthropic client
    const anthropic = this.anthropicService.getClient();
    
    // Use the advanced model for multimodal analysis
    const modelToUse = this.anthropicService.getModel(ModelType.Advanced);
    
    // Make the API call
    return anthropic.messages.create({
      model: modelToUse,
      max_tokens: 4000,
      temperature: 0.2, // Lower temperature for more precise analysis
      system: systemPrompt,
      messages: [
        {
          role: 'user',
          content: messageContent
        }
      ]
    });
  }
  
  /**
   * Extracts text from an Anthropic API response
   * @param response The response from the Anthropic API
   * @returns The extracted text
   */
  private extractTextFromResponse(response: any): string {
    return response.content
      .filter((item: any) => item.type === 'text')
      .map((item: any) => (item.type === 'text' ? item.text : ''))
      .join('');
  }
}
