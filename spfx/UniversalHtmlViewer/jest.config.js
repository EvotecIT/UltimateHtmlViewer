module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jsdom',
  roots: ['<rootDir>/src'],
  testMatch: ['**/__tests__/**/*.[jt]s?(x)'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  moduleNameMapper: {
    '\\.(css|scss)$': '<rootDir>/test/styleMock.js',
    '\\.resx$': '<rootDir>/test/resxMock.js',
  },
  collectCoverage: true,
  collectCoverageFrom: [
    '<rootDir>/src/webparts/universalHtmlViewer/**/*.{ts,tsx}',
    '!<rootDir>/src/webparts/universalHtmlViewer/**/*.d.ts',
    '!<rootDir>/src/webparts/universalHtmlViewer/**/__tests__/**',
  ],
  coverageReporters: ['text-summary', 'lcov'],
  coverageThreshold: {
    global: {
      statements: 18,
      branches: 13,
      functions: 20,
      lines: 18,
    },
  },
};
