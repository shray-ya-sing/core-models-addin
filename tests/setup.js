// setup.js
// This file is also loaded when running tests

// Setup mock directories needed for tests
if (typeof window !== 'undefined') {
  // Mock the Office object for client-side tests
  window.Office = global.Office;
}

// Create file mock for imports of non-JS/TS files
jest.mock('fs', () => ({
  readFileSync: jest.fn().mockImplementation((path) => {
    if (path.endsWith('.pdf')) {
      return Buffer.from('Mock PDF content');
    }
    return 'Mock file content';
  }),
  promises: {
    readFile: jest.fn().mockResolvedValue('Mock file content'),
  },
}));

// Mock any specific modules that cause problems during testing
jest.mock('@anthropic-ai/sdk', () => {
  return {
    Anthropic: jest.fn().mockImplementation(() => ({
      messages: {
        create: jest.fn().mockResolvedValue({
          content: [{ type: 'text', text: 'Mock response' }],
        }),
      },
    })),
  };
});
