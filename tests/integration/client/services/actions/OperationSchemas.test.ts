import { expect } from 'chai';
// Add OpenAI shim for Node.js environment
import 'openai/shims/node';
import OpenAI from 'openai';
import { zodToJsonSchema } from 'zod-to-json-schema';
import { zodTextFormat } from 'openai/helpers/zod';
import { 
  excelCommandPlanSchema,
  excelCommandPlanSchemaJSON,
  openAICompatibleCommandPlanSchema,
  detailedOpenAICommandPlanSchema,
  finalOpenAICommandPlanSchema,
  excelOperationFunctions,
  excelOperationTools
} from '../../../../../src/client/services/actions/OperationSchemas';

describe('OperationSchemas OpenAI Validation', () => {
  
  it('should convert to a valid OpenAI-compatible JSON schema', () => {
    // First, test conversion with a custom name
    const jsonSchemaWithName = zodToJsonSchema(excelCommandPlanSchema, 'command_plan');
    
    // Then, test conversion without a name
    const jsonSchemaWithoutName = zodToJsonSchema(excelCommandPlanSchema);
    
    // Basic validation of the schema structure
    expect(jsonSchemaWithName).to.be.an('object');
    
    // Log schema structure for debugging
    console.log('JSON Schema structure:', JSON.stringify(jsonSchemaWithName).slice(0, 300) + '...');
    
    // Test OpenAI specific format
    try {
      // Convert to OpenAI's expected format WITHOUT a custom name
      const openAIFormat = zodTextFormat(excelCommandPlanSchema, 'command_plan');
      
      // Log the converted schema
      console.log('OpenAI Schema Format:', JSON.stringify(openAIFormat).slice(0, 300) + '...');
      

      // Verify it's serializable and parseable JSON
      const serialized = JSON.stringify(openAIFormat);
      const parsed = JSON.parse(serialized);
      expect(parsed).to.deep.equal(openAIFormat);
      
      // Validate a sample plan with the schema
      const validPlan = {
        description: 'Test command plan',
        operations: [
          {
            op: 'set_value',
            target: 'Sheet1!A1',
            value: 'test value'
          },
          {
            op: 'export_to_pdf',
            sheet: 'Sheet1',
            fileName: 'export',
            quality: null,
            orientation: 'landscape'
          },
          {
            op: 'set_print_settings',
            sheet: 'Sheet1',
            orientation: null,
            printErrors: null,
            printTitles: null
          }
        ]
      };
      
      const result = excelCommandPlanSchema.safeParse(validPlan);
      expect(result.success).to.be.true;
      
      
    }
    catch (error) {
      console.error('Error validating plan:', error);
      throw error; // Re-throw to fail the test
    }
  });
});


describe('OperationSchemas OpenAI Call', () => {
  
  it('should return a valid response from openai', async () => {
    // First, test conversion with a custom name
    // Check if we have an OpenAI API key in the environment for testing
    const openAIFormat = zodTextFormat(excelCommandPlanSchema, 'command_plan');
    const apiKey = '';
    const simpleFormat = {
      name: "command_plan",
      type: "json_schema",
      strict: true,
      schema: {
        type: "object",
        properties: {
          description: { type: "string" },
          operations: { 
            type: "array",
            items: {
              type: "object",
              properties: {
                op: {
                  type: "string" 
                },
                target: {
                  type: "string" 
                },
                value: {
                  type: "string" 
                }
              },
              required: ["op", "target", "value"],
              additionalProperties: false
            }
          }
        },
        required: ["description", "operations"],
        additionalProperties: false
      }
    };

    if (apiKey) {
      // Create an OpenAI instance
      const openai = new OpenAI({
        apiKey: apiKey,
        dangerouslyAllowBrowser: true
      });

      // check that excel operation tools is not empty
      expect(excelOperationTools.length).to.be.greaterThan(0);
      if (excelOperationTools.length === 0) {
        throw new Error("Excel operation tools is empty");
      }
      
      // Create a mock request format for testing
      const requestOptions: any = {
        model: "ft:gpt-4.1-nano-2025-04-14:personal:op:BVoH0tMZ:ckpt-step-26", // Use a fine-tuned model for testing
        input: [
          {
            role: "system",
            content: "You are an Excel assistant that generates operations."
          },
          {
            role: "user",
            content: "Create an Analysis at Various Prices in a new tab. Make the price range from $75 to $80 in increments of $1"
          }
        ],
        max_output_tokens: 1000,
        temperature: 0.1,
        top_p: 1,
        tools: excelOperationTools,
        tool_choice: "required",        
        text: {
          format: simpleFormat
        }
      };
      
      // Since this is an integration test that requires an API key,
      // we'll make it conditional and just verify the request format
      console.log('\nOpenAI Request Format:', JSON.stringify(requestOptions, null, 2).slice(0, 500) + '...');
      
      // If the API key exists, we can attempt to make a real API call
      // This section is commented out as it requires a valid API key and costs money
      // Uncomment for actual API testing:
      try {
        const response = await openai.responses.parse(requestOptions);
        console.log('OpenAI response:', response);
        expect(response).to.be.an('object');
      } catch (error) {
        console.error('OpenAI API Error:', error);
        throw error;
      }
      
    } else {
      console.log('\nSkipping OpenAI API test - No API key available');
    }
  });
});

describe('OperationSchemas OpenAI Function Calling', () => {
  jest.setTimeout(60000);
  
  it('should return valid function calls from OpenAI', async () => {
    const apiKey = '';

    if (apiKey) {
      // Create an OpenAI instance
      const openai = new OpenAI({
        apiKey: apiKey,
        dangerouslyAllowBrowser: true
      });

      const tools = [{
        type: "function",
        name: "add_formula",
        description: "Add a formula to a cell or range",
        parameters: {
          type: "object",
          properties: {
              target: { type: "string" },
              formula: { type: "string" }
          },
          required: ["target", "formula"],
          additionalProperties: false
        },
        strict: true
      },
      {
        type: "function",
        name: "set_value",
        description: "Set a value in a cell or range",
        parameters: {
          type: "object",
          properties: {
            target: {
              type: "string"
            },
            value: {
              type: ["string", "number", "boolean"]
            }
          },
          required: ["target", "value"],
          additionalProperties: false
        }
      }];
      
      // Create request options with function calling
      const requestOptions: any = {
        model: "gpt-4o-mini",
        input: [
          {
            role: "system",
            content: "You are an Excel assistant that generates operations."
          },
          {
            role: "user",
            content: "Set the value of cell A1 to 'Header' and add a formula in B1 that multiplies A1 by 2."
          }
        ],
        max_output_tokens: 300,
        temperature: 0.1,
        top_p: 1,
        tools: tools,
        tool_choice: "required",    
        stream: true,
        store: true
      };
      
      console.log('\nOpenAI Function Calling Request:', JSON.stringify(requestOptions, null, 2).slice(0, 500) + '...');
      
      try {
        const stream = await openai.responses.create(requestOptions) as unknown as AsyncIterable<any>;
        const events = [];
        for await (const event of stream) {
          console.log(event)

        }

        
        // Check that we have function calls
        // select items from response.output that have a type function_call
        const toolCalls = stream.output.filter((item) => item.type === 'function_call');
        expect(toolCalls.length).to.be.at.least(1);
        
        // Validate each tool call
        for (const toolCall of toolCalls) {
          expect(toolCall).to.be.an('object');
          expect(toolCall.name).to.be.a('string');
          expect(['set_value', 'add_formula']).to.include(toolCall.name);
          
          const args = JSON.parse(toolCall.arguments);
          
          if (toolCall.name === 'set_value') {
            expect(args).to.have.property('target');
            expect(args).to.have.property('value');
          } else if (toolCall.name === 'add_formula') {
            expect(args).to.have.property('target');
            expect(args).to.have.property('formula');
          }
        }
        
      } catch (error) {
        console.error('OpenAI API Error:', error);
        throw error;
      }
    } else {
      console.log('\nSkipping OpenAI Function Calling test - No API key available');
    }
  });
});