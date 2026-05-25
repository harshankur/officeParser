import { ConversionResult, GeneratorConfig, OfficeParserAST, OfficeWarningType } from '../types.js';
import { isBrowser } from '../utils/envUtils.js';
import { getAbortError } from '../utils/errorUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { HtmlGenerator } from './HtmlGenerator.js';

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

    async generate(): Promise<ConversionResult> {

        // Step 1: Generate high-fidelity HTML as the source for PDF rendering
        // We reuse the current configuration but ensure standalone mode is on for HTML
        const htmlGenerator = new HtmlGenerator(this.ast, {
            ...this.config,
            htmlConfig: { ...this.config.htmlConfig, standalone: true },
            mdConfig: undefined,
            chunksConfig: undefined,
            csvConfig: undefined,
            textConfig: undefined,
            pdfConfig: undefined,
            rtfConfig: undefined,
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
    private async generateInNode(html: string): Promise<ConversionResult> {
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
            // Dynamic import for peer dependency
            // @ts-ignore
            const puppeteerModule = await import('puppeteer');
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

            const pdfConfig = this.config.pdfConfig;
            const timeout = pdfConfig.timeout;
            if (timeout !== undefined && timeout > 0) {
                page.setDefaultTimeout(timeout);
                page.setDefaultNavigationTimeout(timeout);
            }

            // Set content and wait for network/assets to load
            await page.setContent(html, { waitUntil: 'networkidle0' });

            const pdfBuffer = await page.pdf({
                format: pdfConfig.format,
                width: pdfConfig.width,
                height: pdfConfig.height,
                landscape: pdfConfig.landscape,
                printBackground: pdfConfig.printBackground,
                scale: pdfConfig.scale,
                margin: pdfConfig.margin,
                displayHeaderFooter: pdfConfig.displayHeaderFooter,
                headerTemplate: pdfConfig.headerTemplate,
                footerTemplate: pdfConfig.footerTemplate,
            });

            await browser.close();
            browser = null;

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
            if (err.message && (err.message.includes('timeout') || err.message.includes('Timeout'))) {
                this.warn(OfficeWarningType.PAGE_LOAD_FAILED, `PDF generation timed out: ${err.message}`);
            } else {
                this.warn(OfficeWarningType.DEPENDENCY_LOAD_FAILED, `puppeteer. Please install it with 'npm install puppeteer'. Error: ${err.message}`);
            }
            return {
                value: new Uint8Array(),
                messages: this.messages
            };
        } finally {
            if (signal) {
                signal.removeEventListener('abort', onAbort);
            }
        }
    }

    /**
     * Browser implementation using hidden iframe and native print.
     */
    private async generateInBrowser(html: string): Promise<ConversionResult> {
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
