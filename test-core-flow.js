// Custom test script to run only the core flow tests
const { execSync } = require('child_process');

// Define specific test files to run
const coreFlowTests = [
  'server/llm/services/QueryProcessorService',
  'server/llm/services/KnowledgeBaseCommandIntegration',
  'server/llm/anthropicService',
  'server/llm/execution/CommandExecutor',
  'server/llm/execution/OperationExecutor',
  'server/llm/execution/WebSocketServer'
];

// Run the tests with Jest
try {
  console.log('Running core flow tests...');
  execSync(`npx jest ${coreFlowTests.join(' ')} --verbose`, { stdio: 'inherit' });
  console.log('Core flow tests completed successfully!');
} catch (error) {
  console.error('Core flow tests failed with error:', error.message);
  process.exit(1);
}
