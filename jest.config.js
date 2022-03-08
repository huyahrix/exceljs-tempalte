module.exports = {
    verbose: true,
    testEnvironment: 'node',
    moduleFileExtensions: [
        'js',
        'json',
        'node',
    ],
    testRegex: '(/__tests__/.*|\\.test)\\.js$',
    testPathIgnorePatterns: [
        'node_modules',
        'dist',
    ],
    coverageDirectory: 'coverage',
    collectCoverageFrom: [
        '**/unit/**/**/*.test.js',
    ],
    coveragePathIgnorePatterns: [
        '/node_modules/'
    ],
    coverageThreshold: {
        global: {
            branches: 100,
            functions: 100,
            lines: 100,
            statements: 100,
        },
    },
};
