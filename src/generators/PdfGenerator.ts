import { ConversionResult, GeneratorConfig, OfficeErrorType, OfficeParserAST, OfficeWarningType } from '../types.js';
import { isBrowser } from '../utils/envUtils.js';
import { getAbortError, getOfficeError } from '../utils/errorUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { HtmlGenerator } from './HtmlGenerator.js';
import { renderNativePdf } from './pdf/nativePdfEngine.js';

/**
 * Generates high-fidelity PDF documents using a headless browser engine.
 * 
 * Uses an environment-aware strategy:
 * - Node.js: Uses Puppeteer (peer dependency) for server-side rendering.
 * - Browser: Leverages native browser print capabilities.
 */
export class PdfGenerator extends BaseGenerator<'pdf'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'pdf'>) {
        super('pdf', ast, config);
    }

    async generate(): Promise<ConversionResult<'pdf'>> {

        // Native engine: lay the AST out directly with pdf-lib, no browser required (works in Node
        // and the browser alike). The default 'html' engine falls through to the browser path below.
        if (this.config.pdfConfig.engine === 'native') {
            if (this.config.abortSignal?.aborted) throw getAbortError();
            try {
                const value = await renderNativePdf(this.ast, this.config, this.effectiveMetadata, (t, i) => this.warn(t, i));
                return { value, messages: this.messages };
            } catch (err: any) {
                if (this.config.abortSignal?.aborted) throw getAbortError();
                if (err?.officeIssue) throw err;
                throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, this.config, err?.message || String(err));
            }
        }

        // Step 1: Generate high-fidelity HTML as the source for PDF rendering
        // We reuse the current configuration but ensure standalone mode is on for HTML
        const htmlGenerator = new HtmlGenerator(this.ast, {
            ...this.config,
            // The HTML is rendered in a real browser (Puppeteer in Node, the print engine in-page), which
            // resolves a data: URI of any size. The maxInlineImageBytes cap exists only to keep a huge
            // base64 line out of Markdown/HTML text consumers; applied here it would degrade a large image
            // to a bare `<img src="name">` with no resolvable URL (a broken image in the PDF). Inline fully.
            maxInlineImageBytes: Infinity,
            // Force sourceAttributes off: those data-* attributes are wire-format plumbing for
            // structured consumers and change the mermaid shape's rendered appearance, neither of
            // which belongs in a printed PDF.
            htmlConfig: { ...this.config.htmlConfig, standalone: true, sourceAttributes: false },
        });


        const htmlResult = await htmlGenerator.generate();
        const html = typeof htmlResult.value === 'string' ? htmlResult.value : '';

        // Step 2: Render to PDF based on environment
        if (isBrowser) {
            return this.generateInBrowser(html);
        } else {
            return this.generateInNode(html);
        }
    }

    /**
     * Node.js implementation using Puppeteer.
     * Uses dynamic import to avoid bundling puppeteer into the library core.
     */
    private async generateInNode(html: string): Promise<ConversionResult<'pdf'>> {
        const signal = this.config.abortSignal;
        if (signal?.aborted) {
            throw getAbortError();
        }

        let browser: any;
        const onAbort = async () => {
            if (browser) {
                try {
                    await browser.close();
                } catch (e) {
                    // ignore
                }
            }
        };

        if (signal) {
            signal.addEventListener('abort', onAbort);
        }

        try {
            // Dynamic import for the optional peer dependency. Give a clear, actionable error when it
            // is missing, so the message does not have to be hardcoded into the generic PDF error.
            let puppeteerModule: any;
            try {
                // @ts-ignore
                puppeteerModule = await import('puppeteer');
            } catch {
                throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, this.config,
                    "the default PDF engine requires the optional peer dependency 'puppeteer' (install it with `npm install puppeteer`), or use pdfConfig.engine: 'native' (which uses pdf-lib and needs no browser)");
            }
            const puppeteer = puppeteerModule.default || puppeteerModule;

            const launchOptions = { ...this.config.pdfConfig.launchOptions };

            // Handle Apple Silicon / Rosetta performance warning and binary detection
            const isMac = process.platform === 'darwin';
            const isX64 = process.arch === 'x64';
            let isRosetta = false;

            if (isMac && isX64) {
                try {
                    const { execSync } = await import('child_process');
                    isRosetta = execSync('sysctl -n hw.optional.arm64', { stdio: 'pipe' }).toString().trim() === '1';
                } catch (e) {
                    // Ignore errors in detection
                }
            }

            if (isRosetta) {
                this.warn(OfficeWarningType.PERFORMANCE_TIP, "You are running on Apple Silicon using an x64 Node.js installation. PDF generation will be significantly faster (avoiding Rosetta translation) if you switch to a native arm64 Node.js version.");
            }

            // Note: We are no longer suppressing the Puppeteer 'Degraded performance' warning here
            // to ensure transparency about the environment state. Programmatically fixing this
            // would require force-downloading a ~300MB arm64 browser binary or switching to 
            // a system-installed Chrome, both of which are too intrusive for a library.

            browser = await puppeteer.launch(launchOptions);
            const page = await browser.newPage();

            // Harden against SSRF: the HTML being rendered is derived from an untrusted
            // document, and `networkidle0` would otherwise fetch every URL it references
            // (external images, stylesheets, etc.) from this host, reaching internal
            // services or a cloud metadata endpoint (169.254.169.254). Intercept requests
            // and allow only inline data/blob URIs and the configured chart CDN; abort every
            // other remote fetch.
            const allowedHosts = new Set<string>();
            try {
                const chartSrc = this.config.htmlConfig?.chartJsSrc;
                if (chartSrc) allowedHosts.add(new URL(chartSrc).host);
            } catch { /* no or invalid chart CDN configured */ }

            let blockedRemoteResource = false;
            await page.setRequestInterception(true);
            page.on('request', (req: any) => {
                const url = req.url();
                if (url.startsWith('data:') || url.startsWith('blob:') || url.startsWith('about:')) {
                    return req.continue().catch(() => { /* request already handled */ });
                }
                try {
                    if (allowedHosts.has(new URL(url).host)) {
                        return req.continue().catch(() => { /* request already handled */ });
                    }
                } catch { /* unparseable URL: fall through and block */ }
                blockedRemoteResource = true;
                return req.abort().catch(() => { /* request already handled */ });
            });

            const pdfConfig = this.config.pdfConfig;
            const timeout = pdfConfig.timeout;
            if (timeout !== undefined && timeout > 0) {
                page.setDefaultTimeout(timeout);
                page.setDefaultNavigationTimeout(timeout);
            }

            // Set content and wait for network/assets to load
            await page.setContent(html, { waitUntil: 'networkidle0' });

            // Margin sides default to the '' unset sentinel; the HTML/Puppeteer default is zero margins
            // (the standalone body carries its own padding), so map unset -> 0 and pass any explicit
            // value through untouched.
            const m = pdfConfig.margin;
            const marginPx = (v: string | number) => (v === '' ? 0 : v);

            const pdfBuffer = await page.pdf({
                format: pdfConfig.format,
                width: pdfConfig.width,
                height: pdfConfig.height,
                landscape: pdfConfig.landscape,
                printBackground: pdfConfig.printBackground,
                scale: pdfConfig.scale,
                margin: { top: marginPx(m.top), right: marginPx(m.right), bottom: marginPx(m.bottom), left: marginPx(m.left) },
                displayHeaderFooter: pdfConfig.displayHeaderFooter,
                headerTemplate: pdfConfig.headerTemplate,
                footerTemplate: pdfConfig.footerTemplate,
                // Emit a tagged (accessible) PDF and, optionally, a bookmark outline from the
                // document headings. Ignored by older Puppeteer versions that lack these options.
                tagged: pdfConfig.tagged,
                outline: pdfConfig.outline,
            });

            await browser.close();
            browser = null;

            if (blockedRemoteResource) {
                this.warn(OfficeWarningType.BROWSER_GENERATION_LIMITATION, 'One or more remote resources referenced by the document were blocked during PDF rendering to prevent server-side request forgery (SSRF). Only inline images and the configured chart CDN are loaded.');
            }

            return {
                value: new Uint8Array(pdfBuffer),
                messages: this.messages
            };
        } catch (err: any) {
            if (browser) {
                try {
                    await browser.close();
                } catch (e) {
                    // ignore
                }
                browser = null;
            }
            if (signal?.aborted) {
                throw getAbortError();
            }
            // Never return an empty PDF: a zero-byte buffer looks like success to callers who don't
            // inspect `messages` and silently produces broken files. Fail loudly with a typed error.
            throw getOfficeError(OfficeErrorType.PDF_GENERATION_FAILED, this.config, err?.message || String(err));
        } finally {
            if (signal) {
                signal.removeEventListener('abort', onAbort);
            }
        }
    }

    /**
     * Browser implementation using hidden iframe and native print.
     */
    private async generateInBrowser(html: string): Promise<ConversionResult<'pdf'>> {
        this.warn(OfficeWarningType.BROWSER_GENERATION_LIMITATION, "Browser-based PDF generation triggered. For automated 'Save as PDF' without user interaction, we recommend using 'html2pdf.js' as a custom generator hook.");

        // In a browser environment, we return the HTML and suggest using window.print()
        // Or we could trigger a print dialog immediately if desired, 
        // but returning the string allows the user to decide where to inject it.
        return {
            value: html,
            messages: this.messages
        };
    }
}
