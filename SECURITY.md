# Security Policy

`officeParser` processes documents that may contain sensitive information and often come from untrusted sources. I welcome vulnerability reports and will address legitimate ones as best I can. Please read the security posture and disclaimer below before assuming any particular guarantee, and note that, as a single maintainer, I make best-effort commitments rather than promises.

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

### On Receiving a Report

I am a single maintainer with no dedicated security team, so I cannot commit to a fixed response or resolution timeline. When I receive a valid report I will do my best to:
*   Acknowledge it when I am able to.
*   Triage it, assess its impact, and validate it with you.
*   Work with you on a fix and ship it once it is ready.
*   Credit you in the release notes (optional, at your discretion).

I ask that you follow **coordinated disclosure** practices and allow a reasonable window for a patch before sharing technical details publicly.

## Security Posture & Disclaimer

`officeParser` is a parser, generator, and converter for office documents, and is often used on
files from untrusted sources (user uploads, email attachments, scraped content). Like any parser
that accepts arbitrary input, it has a large attack surface. I do sanitize output and apply
hardening where I can (injection escaping, decompression limits, some resource and recursion
bounds, SSRF precautions during PDF rendering) and I fix issues as I find them, but this is
best-effort, not a guarantee: a library of this size will have vectors I have not found or have
not yet addressed.

Treat it as garbage in, garbage out. Sanitize and validate untrusted files at your own boundary,
and run parsing in isolation appropriate to your threat model (sandboxing or containerization,
memory and time limits, a low-privilege process, and the `abortSignal` and `decompressionLimits`
options this library exposes) rather than relying on the library's hardening alone. No software
can be guaranteed free of vulnerabilities; `officeParser` is provided "AS IS", without warranty of
any kind, per the [LICENSE](LICENSE). Final responsibility for the impact of a malicious file on
your system rests with you, the consumer of the library.

## Out of Scope
- **Third-Party Dependencies**: Vulnerabilities in upstream libraries (e.g., `tesseract.js`, `pdfjs-dist`, `fflate`) should be reported directly to their respective maintainers.
- **Semantic Chunking & External APIs**: When using `Semantic Strategy` for chunking, you provide an `embed` function. If this function sends data to external LLM providers (like OpenAI or Anthropic), the security and privacy of that data transfer are governed by your implementation and the provider's terms. `officeParser` itself does not transmit document data to external servers by default.
