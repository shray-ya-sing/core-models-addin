// setupTests.js
// This file is run before each test file

// Import the necessary testing libraries
import '@testing-library/jest-dom';

// Mock Office objects globally
global.Office = {
  context: {
    requirements: {
      isSetSupported: jest.fn().mockReturnValue(true)
    }
  },
  initialize: jest.fn(callback => callback())
};

// Mock console methods to avoid cluttering test output
// while still allowing errors to be seen
const originalConsoleError = console.error;
const originalConsoleWarn = console.warn;

// Preserve console.error for actual errors but suppress expected warnings
console.error = (...args) => {
  if (args[0] && typeof args[0] === 'string' && 
      (args[0].includes('test was not wrapped in act') || 
       args[0].includes('validateDOMNesting'))) {
    return;
  }
  originalConsoleError(...args);
};

// Suppress specific warning messages
console.warn = (...args) => {
  if (args[0] && typeof args[0] === 'string' && 
      (args[0].includes('componentWillReceiveProps') || 
       args[0].includes('componentWillUpdate') ||
       args[0].includes('componentWillMount'))) {
    return;
  }
  originalConsoleWarn(...args);
};

// Setup fetch mock
global.fetch = jest.fn();

// Reset mocks between tests
beforeEach(() => {
  jest.clearAllMocks();
});
