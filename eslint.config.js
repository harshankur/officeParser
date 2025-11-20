const js = require('@eslint/js');
const globals = require('globals');

module.exports = [
    js.configs.recommended,
    {
        files: ['**/*.js'],
        languageOptions: {
            ecmaVersion: 2020,
            sourceType: 'commonjs',
            globals: {
                ...globals.node,
                ...globals.es2020,
            },
        },
        rules: {
            // Possible Problems
            'no-unused-vars': ['warn', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
            'no-undef': 'error',
            'no-constant-condition': 'warn',
            'no-duplicate-imports': 'error',

            // Best Practices
            'eqeqeq': ['error', 'always', { null: 'ignore' }],
            'no-var': 'warn',
            'prefer-const': 'warn',
            'no-eval': 'error',
            'no-implied-eval': 'error',
            'no-console': 'off',

            // Stylistic (deprecated but still available)
            'semi': ['error', 'always'],
            'quotes': ['error', 'single', { avoidEscape: true, allowTemplateLiterals: true }],
            'indent': ['error', 4, { SwitchCase: 1 }],
            'comma-dangle': ['error', 'only-multiline'],
            'no-trailing-spaces': 'error',
            'eol-last': ['error', 'always'],
            'no-multiple-empty-lines': ['error', { max: 2, maxEOF: 0 }],
            'brace-style': ['error', '1tbs', { allowSingleLine: true }],
            'comma-spacing': ['error', { before: false, after: true }],
            'keyword-spacing': ['error', { before: true, after: true }],
            'space-before-blocks': 'error',
            'space-infix-ops': 'error',
        },
    },
    {
        ignores: [
            'node_modules/**',
            'typings/**',
            'pdfjs-dist-build/**',
            'test/**',
            '.idea/**',
            'simons_test/**',
        ],
    },
];
