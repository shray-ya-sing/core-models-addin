/**
 * Mock implementation of Anthropic API for unit testing
 */

export const mockAnthropicClient = {
  messages: {
    create: jest.fn().mockResolvedValue({
      id: "msg_mock123456",
      type: "message",
      role: "assistant",
      content: [
        {
          type: "text",
          text: '{"description":"Test operations","operations":[]}'
        }
      ],
      model: "claude-3-haiku-20240307",
      usage: {
        input_tokens: 150,
        output_tokens: 50
      }
    })
  }
};

export class MockClientAnthropicService {
  getClient() {
    return mockAnthropicClient;
  }

  getModel() {
    return "claude-3-haiku-20240307";
  }

  extractJsonFromMarkdown(markdown: string): string {
    // Simple mock implementation that extracts JSON between code blocks
    const jsonMatch = markdown.match(/```json\n([\s\S]*?)\n```/) || 
                     markdown.match(/```\n([\s\S]*?)\n```/) || 
                     [null, markdown];
    return jsonMatch[1] || markdown;
  }
}
