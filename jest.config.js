module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jsdom',
  moduleNameMapper: {
    '\\.(css|less|scss|sass)$': 'identity-obj-proxy',
    '\\.(jpg|jpeg|png|gif|eot|otf|webp|svg|ttf|woff|woff2|mp4|webm|wav|mp3|m4a|aac|oga)$': '<rootDir>/tests/__mocks__/fileMock.js',
    '^@anthropic-ai/sdk$': '<rootDir>/tests/__mocks__/@anthropic-ai/sdk',
    '^pdf-parse$': '<rootDir>/tests/__mocks__/pdf-parse',
  },
  setupFilesAfterEnv: ['<rootDir>/tests/setupTests.js', '<rootDir>/tests/setup.js'],
  testPathIgnorePatterns: ['/node_modules/'],
  transform: {
    '^.+\\.(ts|tsx)$': ['ts-jest', {
      tsconfig: 'tsconfig.test.json'
    }],
    '^.+\\.(js|jsx)$': 'babel-jest',
  },
  transformIgnorePatterns: [
    '/node_modules/(?!(@fluentui|uuid|@xenova)/)',
  ],
  moduleDirectories: ['node_modules', 'src'],
  testTimeout: 30000,
};