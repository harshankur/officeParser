# Security Policy

I am committed to the security and privacy of the users of this library. As a library that processes documents containing potentially sensitive information, I treat all vulnerability reports with the highest priority.

## Supported Versions

Maintenance and security patches are currently provided for the latest major branch.

| Version | Supported          |
| ------- | ------------------ |
| v6.x    | :white_check_mark: |
| v5.x    | :x:                |
| < v5.x  | :x:                |

## Reporting a Vulnerability

**Please do not open public issues for security vulnerabilities.**

### 🛑 Why Private Disclosure?
Publicly disclosing a vulnerability before a fix is available creates an immediate window for attackers to exploit the issue against the millions of users currently running this library. Private reporting allows us to develop, test, and release a patch **before** the technical details are shared, ensuring that the community remains protected.

To facilitate a safe disclosure, please use our private GitHub reporting channel:

### GitHub Private Reporting
1.  Navigate to the [Security tab](https://github.com/harshankur/officeParser/security) of this repository.
2.  Select **Advisories** from the left-hand navigation.
3.  Click the **Report a vulnerability** button.
    *   *Note: This creates a secure, private thread between you and the maintainer to collaborate on a fix.*

### My Security Commitment

Upon receiving a valid report, I will:
*   Acknowledge receipt within **48 hours**.
*   Perform a thorough triage and impact assessment within **5 business days**.
*   Work with the reporter to validate and test the fix.
*   Credit the researcher in the release notes (optional, at your discretion).

I ask that you follow **coordinated disclosure** practices and allow a reasonable window to release a patch before sharing technical details publicly.

## Out of Scope
Vulnerabilities in third-party dependencies (e.g., `tesseract.js`, `pdfjs-dist`) should be reported directly to their respective upstream maintainers.
