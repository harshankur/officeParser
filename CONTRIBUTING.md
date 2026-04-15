# Contributing to officeParser 🚀

First off, thank you for considering a contribution to `officeParser`! Maintaining a project that handles over 260,000 weekly installs is a significant responsibility, and your help is vital for its long-term stability.

## Code of Conduct
This project is governed by our [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected to uphold these standards.

## How to Contribute

### 1. Reporting Bugs 🐛
We use interactive **GitHub Issue Forms** to help categorize bugs by internal library component. 
- Please use the [Bug Report Form](https://github.com/harshankur/officeParser/issues/new?template=bug_report.yml).
- **Pro Tip**: Providing a sample file (via the upload field) is the single fastest way to get a bug fixed.

### 2. Suggesting Enhancements 💡
Technical suggestions are welcome! Please use the [Feature Request Form](https://github.com/harshankur/officeParser/issues/new?template=feature_request.yml) to outline your goal and the expected impact on the library.

### 3. Pull Requests 🛠️
We welcome pull requests that improve performance, expand format support, or fix documented bugs.

**The PR Process:**
1.  Fork the repo and create your branch from `master`.
2.  **Local Testing**: Run `npm test` to ensure all existing parsers (Word, PDF, Excel, etc.) still pass their baseline checks.
3.  **Build Validation**: Run `npm run build` to ensure the TypeScript-to-JS compilation and browser bundling function correctly.
4.  **Documentation**: If you've changed the AST structure or added a new config option, please update relevant documentation or types.
5.  Submit your PR!

## Development Setup

```bash
# 1. Clone the repository
git clone https://github.com/harshankur/officeParser.git
cd officeParser

# 2. Install dependencies
npm install

# 3. Running the standard test suite
npm test

# 4. Building the project (ESM, CJS, and Browser bundles)
npm run build
```

## Coding Standards
- **Strict Typing**: All new code must be strictly typed. Avoid `any` at all costs.
- **AST Integrity**: Ensure your changes do not break the "Hierarchical AST" philosophy of the library.
- **Zero-Dependency Core**: My goal is to keep the core dependency list extremely lean. Propose new dependencies in an issue before adding them to a PR.

## Questions?
If you have a general question about how to use the library or its internal architecture, please open a [Discussion](https://github.com/harshankur/officeParser/discussions) rather than an issue.
