import { OfficeErrorType, OfficeParserConfig } from '../types.js';
import { checkAbortSignal, getOfficeError } from '../utils/errorUtils.js';
import { isCfb, isEncryptedOoxml, decryptOoxml, isEncryptedOdf, decryptOdf, WRONG_PASSWORD } from './index.js';

/** How many times an `onPassword` callback may be re-asked before decryption gives up. */
export const MAX_PASSWORD_ATTEMPTS = 3;

/**
 * Decrypts a password-protected OOXML or ODF container, returning the plaintext bytes; a buffer that
 * is not an encrypted container is returned unchanged. Uses the top-level `password`, then retries
 * through `onPassword` (capped at {@link MAX_PASSWORD_ATTEMPTS}), and rejects with
 * `PASSWORD_REQUIRED`/`PASSWORD_INCORRECT` exactly as the PDF path does. PDF encryption is not handled
 * here (pdf.js does it inside the PDF parser).
 *
 * Shared by the parser and the template renderer so both honour `password` and `onPassword`
 * identically, and neither drifts from the other's retry/limit semantics.
 */
export const decryptIfNeeded = async (buffer: Buffer, config: OfficeParserConfig, extHint?: string): Promise<Buffer> => {
    checkAbortSignal(config.abortSignal);
    const limits = config.decompressionLimits;
    let decrypt: ((buf: Uint8Array, password: string) => Uint8Array | Promise<Uint8Array>) | null = null;
    if (isCfb(buffer)) {
        // A CFB container is either an encrypted OOXML file or a legacy binary (.doc/.xls/.ppt).
        // Only the former carries the encryption streams; the latter falls through as unsupported.
        if (isEncryptedOoxml(buffer)) decrypt = decryptOoxml;
    } else if (await isEncryptedOdf(buffer, limits, extHint, config.abortSignal)) {
        // The ODF detection and decryption reuse the caller's decompression limits, so a zip-bomb
        // in a would-be encrypted ODF is bounded exactly as an ordinary document is. The PBKDF2 budget
        // is created here, outside the retry loop below, so the whole password dance shares one
        // allowance instead of granting a hostile file a fresh one on every attempt.
        const budget = { spent: 0 };
        decrypt = (b, password) => decryptOdf(b, password, limits, config, budget);
    }
    checkAbortSignal(config.abortSignal);
    if (!decrypt) return buffer;

    const onPassword = config.onPassword;
    let password = config.password || '';
    let attempts = 0;
    while (true) {
        checkAbortSignal(config.abortSignal);
        if (!password) {
            if (onPassword && attempts < MAX_PASSWORD_ATTEMPTS) {
                attempts++;
                const supplied = await onPassword('required');
                checkAbortSignal(config.abortSignal);
                if (supplied) { password = supplied; continue; }
            }
            throw getOfficeError(OfficeErrorType.PASSWORD_REQUIRED, config);
        }
        try {
            const decrypted = await decrypt(buffer, password);
            checkAbortSignal(config.abortSignal);
            return Buffer.from(decrypted);
        } catch (e) {
            // Cancellation is not a decryption failure and must not trigger a password retry.
            if (typeof e === 'object' && e !== null && 'name' in e && e.name === 'AbortError') throw e;
            checkAbortSignal(config.abortSignal);
            if (e !== WRONG_PASSWORD) {
                // A decompression-limit breach thrown by extractFiles during decryption already carries a
                // typed officeIssue (and was reported once): re-throw it unchanged rather than masking its
                // code and reporting it a second time.
                if ((e as { officeIssue?: unknown })?.officeIssue) throw e;
                // Any other structural/unsupported-scheme failure (not a wrong password): surface it as a
                // typed decryption error rather than a raw throw.
                throw getOfficeError(OfficeErrorType.DOCUMENT_DECRYPTION_FAILED, config, e instanceof Error ? e.message : String(e));
            }
            password = '';
            if (onPassword && attempts < MAX_PASSWORD_ATTEMPTS) {
                attempts++;
                const supplied = await onPassword('incorrect');
                checkAbortSignal(config.abortSignal);
                if (supplied) { password = supplied; continue; }
            }
            throw getOfficeError(OfficeErrorType.PASSWORD_INCORRECT, config);
        }
    }
};
