# Contributing to officeParser

First off, thanks for taking the time to contribute! 🎉

The following is a set of guidelines for contributing to `officeParser`. These are mostly guidelines, not rules. Use your best judgment, and feel free to propose changes to this document in a pull request.

## Code of Conduct

This project and everyone participating in it is governed by the [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected to uphold this code.

## How Can I Contribute?

### Reporting Bugs

This section guides you through submitting a bug report for `officeParser`. Following these guidelines helps maintainers and the community understand your report, reproduce the behavior, and find related reports.

- **Use a clear and descriptive title** for the issue to identify the problem.
- **Describe the exact steps to reproduce the problem** in as much detail as possible.
- **Provide specific examples** to demonstrate the steps.
- **Describe the behavior you observed after following the steps** and point out what exactly is the problem with that behavior.
- **Explain which behavior you expected to see instead and why.**
- **Include snippets of the file** that is causing the issue, or attach a small sample file if possible (ensure no private data involved).

### Suggesting Enhancements

This section guides you through submitting an enhancement suggestion for `officeParser`, including completely new features and minor improvements to existing functionality.

- **Use a clear and descriptive title** for the issue to identify the suggestion.
- **Provide a step-by-step description of the suggested enhancement** in as much detail as possible.
- **Explain why this enhancement would be useful** to most `officeParser` users.

### key_v2 Pull Requests

1. Fork the repo and create your branch from `main`.
2. If you've added code that should be tested, add tests.
3. If you've changed APIs, update the documentation.
4. Ensure the test suite passes (`npm test`).
5. Make sure your code lints.
6. Issue that pull request!

## development Setup

1.  **Clone the repository**
    ```bash
    git clone https://github.com/harshankur/officeParser.git
    cd officeParser
    ```

2.  **Install dependencies**
    ```bash
    npm install
    ```

3.  **Build the project**
    ```bash
    npm run build
    ```

4.  **Run tests**
    ```bash
    npm test
    ```

## Coding Style

- We use **TypeScript** for type safety. Please ensure strict mode is enabled (default in `tsconfig.json`).
- Follow the existing code style (indentation, variable naming).
- Run `npm run clean` before submitting to ensure a fresh build.

## Questions?

Feel free to open an issue with the label `question` if you have any doubts.
