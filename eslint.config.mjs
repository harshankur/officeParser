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
            "no-console": "error"
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
            "no-console": "off"
        }
    }
];
