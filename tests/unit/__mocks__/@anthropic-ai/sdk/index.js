// Mock for @anthropic-ai/sdk
const Anthropic = jest.fn().mockImplementation(() => ({
  messages: {
    create: jest.fn().mockResolvedValue({
      content: [{ type: 'text', text: 'Mock response' }],
    }),
  },
}));

module.exports = {
  Anthropic,
  default: Anthropic
};
