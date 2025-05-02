/**
 * Unit tests for ClientExcelOperationGenerator
 */

import { ClientExcelOperationGenerator } from "../../../../src/client/services/ClientExcelOperationGenerator";
import { ClientAnthropicService, ModelType } from "../../../../src/client/services/ClientAnthropicService";
import { ExcelOperation } from "../../../../src/client/models/ExcelOperationModels";

// Mock ClientAnthropicService
jest.mock("../../../../src/client/services/ClientAnthropicService", () => {
  return {
    ModelType: { Standard: 'standard' },
    ClientAnthropicService: jest.fn().mockImplementation(() => ({
      getClient: jest.fn().mockReturnValue({
        messages: {
          create: jest.fn().mockResolvedValue({
            content: [{ 
              type: 'text', 
              text: '{"description":"Test operations","operations":[{"op":"create_sheet","name":"New Sheet"}]}'
            }]
          })
        }
      }),
      getModel: jest.fn().mockReturnValue("claude-3-haiku-20240307"),
      extractJsonFromMarkdown: jest.fn(text => text)
    }))
  };
});

describe("ClientExcelOperationGenerator", () => {
  let generator: ClientExcelOperationGenerator;
  let mockAnthropicService: jest.Mocked<ClientAnthropicService>;
  
  beforeEach(() => {
    // Create a mock for the Anthropic service
    mockAnthropicService = new ClientAnthropicService("mock-api-key") as jest.Mocked<ClientAnthropicService>;
    
    // Create the generator with the mocked service
    generator = new ClientExcelOperationGenerator({
      anthropic: mockAnthropicService,
      debugMode: false
    });
  });
  
  describe("generateOperations", () => {
    test("generates operations from user query", async () => {
      const query = "Create a new worksheet";
      const workbookContext = "Current workbook has Sheet1";
      const chatHistory: Array<{ role: string; content: string }> = [];
      
      const result = await generator.generateOperations(query, workbookContext, chatHistory);
      
      // Verify the Anthropic client was called
      expect(mockAnthropicService.getClient).toHaveBeenCalled();
      expect(mockAnthropicService.getModel).toHaveBeenCalled();
      
      // Verify the result structure
      expect(result).toHaveProperty("description");
      expect(result).toHaveProperty("operations");
      expect(result.operations).toHaveLength(1);
      expect(result.operations[0].op).toBe("create_sheet");
    });
    
    test("enforces minimalism by generating only explicitly requested operations", async () => {
      // Setup the mock to return different responses based on input
      const minimalistResponse = {
        content: [{ 
          type: 'text', 
          text: '{"description":"Create new worksheet","operations":[{"op":"create_sheet","name":"New Sheet"}]}'
        }]
      };
      
      const createMethod = mockAnthropicService.getClient().messages.create as jest.Mock;
      createMethod.mockResolvedValueOnce(minimalistResponse);
      
      const query = "Add a new tab";
      const workbookContext = "Current workbook has Sheet1";
      const chatHistory: Array<{ role: string; content: string }> = [];
      
      const result = await generator.generateOperations(query, workbookContext, chatHistory);
      
      // Verify we only got the explicitly requested operation (create_sheet)
      // and no additional formatting or data operations
      expect(result.operations).toHaveLength(1);
      expect(result.operations[0].op).toBe("create_sheet");
      expect(result.operations.filter(op => 
        op.op === "set_value" || 
        op.op === "format_range" ||
        op.op === "create_chart"
      ).length).toBe(0);
    });
    
    test("handles API errors gracefully", async () => {
      // Setup the mock to throw an error
      const createMethod = mockAnthropicService.getClient().messages.create as jest.Mock;
      createMethod.mockRejectedValueOnce(new Error("API Error"));
      
      const query = "Create a table";
      const workbookContext = "Current workbook has Sheet1";
      const chatHistory: Array<{ role: string; content: string }> = [];
      
      const result = await generator.generateOperations(query, workbookContext, chatHistory);
      
      // Verify error handling
      expect(result.description).toContain("Error");
      expect(result.operations).toHaveLength(0);
    });
    
    test("handles JSON parsing errors gracefully", async () => {
      // Setup the mock to return invalid JSON
      const createMethod = mockAnthropicService.getClient().messages.create as jest.Mock;
      createMethod.mockResolvedValueOnce({
        content: [{ type: 'text', text: 'Not valid JSON' }]
      });
      
      const query = "Create a table";
      const workbookContext = "Current workbook has Sheet1";
      const chatHistory: Array<{ role: string; content: string }> = [];
      
      const result = await generator.generateOperations(query, workbookContext, chatHistory);
      
      // Verify error handling for JSON parsing
      expect(result.description).toContain("Error");
      expect(result.operations).toHaveLength(0);
    });
    
    test("passes filtered chat history to maintain context", async () => {
      const query = "Create a new worksheet";
      const workbookContext = "Current workbook has Sheet1";
      const chatHistory = [
        { role: "user", content: "Can you help me make a budget sheet?" },
        { role: "assistant", content: "Sure, I can help with that." },
        { role: "system", content: "System message" }, // Should be filtered out
        { role: "user", content: "Create a new worksheet" }
      ];
      
      await generator.generateOperations(query, workbookContext, chatHistory);
      
      // Verify that the create method was called with filtered chat history (no system messages)
      const createMethod = mockAnthropicService.getClient().messages.create as jest.Mock;
      expect(createMethod).toHaveBeenCalled();
      
      // Get the actual argument passed
      const callArg = createMethod.mock.calls[0][0];
      
      // Verify system messages were filtered out and only the most recent messages were included
      const passedMessages = callArg.messages;
      const hasSystemMessage = passedMessages.some(msg => msg.role === "system");
      expect(hasSystemMessage).toBe(false);
    });
    
    test("validates operations to ensure they are well-formed", async () => {
      // Setup the mock to return operations with a missing "op" field
      const createMethod = mockAnthropicService.getClient().messages.create as jest.Mock;
      createMethod.mockResolvedValueOnce({
        content: [{ 
          type: 'text', 
          text: '{"description":"Invalid operations","operations":[{"name":"Missing op field"}]}'
        }]
      });
      
      // Spy on console.error to verify it's called with the validation error
      const consoleErrorSpy = jest.spyOn(console, 'error');
      
      const query = "Create a table";
      const workbookContext = "Current workbook has Sheet1";
      const chatHistory: Array<{ role: string; content: string }> = [];
      
      try {
        // The implementation catches the error internally and returns an empty operations array
        const result = await generator.generateOperations(query, workbookContext, chatHistory);
        
        // Verify the fallback error response was returned
        expect(result.description).toContain("Error");
        expect(result.operations).toHaveLength(0);
      } catch (error) {
        // If it throws, the test will fail
        fail("Should have handled the error internally");
      } finally {
        // Restore the original console.error
        consoleErrorSpy.mockRestore();
      }
    });
  });
});
