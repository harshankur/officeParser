# Security Policy

I am committed to the security and privacy of the users of this library. As a library that processes documents containing potentially sensitive information, I treat all vulnerability reports with the highest priority.

## Supported Versions

Maintenance and security patches are currently provided for the latest major branch.

| Version | Supported          |
| ------- | ------------------ |
| v7.x    | :white_check_mark: |
| v6.x    | :white_check_mark: |
| < v6.x  | :x:                |

## Reporting a Vulnerability

**Please do not open public issues for security vulnerabilities.**

### 🛑 Why Private Disclosure?
Publicly disclosing a vulnerability before a fix is available creates an immediate window for attackers to exploit the issue against the millions of users currently running this library. Private reporting allows me to develop, test, and release a patch **before** the technical details are shared, ensuring that the community remains protected.

To facilitate a safe disclosure, please use my private GitHub reporting channel:

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

## Security Posture & Disclaimer

`officeParser` is a parser, generator, and converter for office documents — it is often used to
process files from untrusted sources (user uploads, email attachments, scraped content, etc.), so
I treat every input as potentially adversarial. I actively invest in hardening the library against
malicious input — output sanitization against injection, zip-bomb and decompression limits,
recursion/resource limits against denial-of-service payloads, SSRF protection during PDF
rendering, and more — and I will keep shipping security fixes for as long as I actively maintain
this project (see CHANGELOG.md for the ongoing history).

That said, I am the sole maintainer of this project, with no dedicated security team behind it,
and **no software can be guaranteed free of vulnerabilities**. I make no warranty that
`officeParser` is or will remain impervious to a sufficiently novel attack — the library is
provided "AS IS", without warranty of any kind, per the [LICENSE](LICENSE). **Final responsibility
for the impact of a compromised or malicious input file on your system rests with you, the
consumer of the library.** If you process files from untrusted sources, apply isolation
appropriate to your own threat model (sandboxing/containerization, resource limits, a
low-privilege execution context, etc.) rather than relying on any single library's hardening as a
complete solution.

## Out of Scope
- **Third-Party Dependencies**: Vulnerabilities in upstream libraries (e.g., `tesseract.js`, `pdfjs-dist`, `fflate`) should be reported directly to their respective maintainers.
- **Semantic Chunking & External APIs**: When using `Semantic Strategy` for chunking, you provide an `embed` function. If this function sends data to external LLM providers (like OpenAI or Anthropic), the security and privacy of that data transfer are governed by your implementation and the provider's terms. `officeParser` itself does not transmit document data to external servers by default.
