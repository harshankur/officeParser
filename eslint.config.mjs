import tsParser from "@typescript-eslint/parser";
import tsPlugin from "@typescript-eslint/eslint-plugin";

export default [
    {
        // Global ignores
        ignores: ["dist/**", "node_modules/**", "docs/**"]
    },
    {
        files: ["**/*.ts"],
        languageOptions: {
            parser: tsParser,
            sourceType: "module",
            ecmaVersion: "latest"
        },
        plugins: {
            "@typescript-eslint": tsPlugin
        },
        rules: {
            // Enforce centralized error reporting by banning direct console usage in the core library
            "no-console": "error",
            "no-restricted-syntax": [
                "error",
                {
                    selector: "CatchClause:not(:has(IfStatement BinaryExpression[operator='==='] Literal[value='AbortError'])) CallExpression[callee.name='getWrappedError']",
                    message: "AbortError must not be passed to getWrappedError. Check for it first in the catch block (e.g., if (error?.name === 'AbortError') throw error;)"
                },
                {
                    selector: "NewExpression[callee.name='Error'][arguments.0.type='Literal'], NewExpression[callee.name='Error'][arguments.0.type='TemplateLiteral']",
                    message: "Do not instantiate Error with a direct string literal or template literal. Use OfficeErrorType enums and the ERROR_MESSAGES dictionary to support future localization/consistency."
                },
                {
                    selector: "NewExpression[callee.name='DOMException'][arguments.0.type='Literal'], NewExpression[callee.name='DOMException'][arguments.0.type='TemplateLiteral']",
                    message: "Do not instantiate DOMException with a direct string literal or template literal. Use OfficeErrorType enums and the ERROR_MESSAGES dictionary to support future localization/consistency."
                }
            ]
        }
    },
    {
        // Exemptions for the CLI and the reporting utility itself
        files: [
            "src/cli.ts",
            "src/utils/errorUtils.ts",
            "test/**/*.ts",
            "scripts/**/*.ts",
            "build_browser.js"
        ],
        rules: {
            "no-console": "off",
            "no-restricted-syntax": "off"
        }
    }
];
