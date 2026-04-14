/**
 * OCR (Optical Character Recognition) Utilities
 * 
 * This module provides functions for extracting text from images using Tesseract.js.
 * Used when `config.ocr` is enabled to extract text from embedded images in documents.
 * 
 * Includes a worker pool via OcrSchedulerManager to improve performance when 
 * processing multiple images.
 * 
 * @module ocrUtils
 */

import { OcrConfig } from '../types.js';

/**
 * Manages a pool of Tesseract workers using a scheduler.
 * This improves performance by reusing workers across multiple images.
 * 
 * Implements lazy loading of tesseract.js to ensure no background processes
 * are spawned unless OCR is explicitly used.
 */
/**
 * Internal interface for tracking workers within the pool.
 */
interface ManagedWorker {
    worker: any;
    language: string;
    lastUsed: number;
    isBusy: boolean;
}

/**
 * Manages a pool of Tesseract workers with "Smart Affinity".
 * 
 * Instead of a simple scheduler, this manager allows workers to persist with 
 * a specific language affinity. If a new language is requested and the pool 
 * is at capacity, it re-initializes the Least Recently Used (LRU) idle worker 
 * rather than resetting the entire pool.
 * 
 * Implements lazy loading of tesseract.js to ensure no background processes
 * are spawned unless OCR is explicitly used.
 */
class OcrSchedulerManager {
    private static instance: OcrSchedulerManager;
    private pool: ManagedWorker[] = [];
    private queue: { image: any, config: OcrConfig, resolve: (text: string) => void, reject: (err: any) => void }[] = [];
    private readonly MAX_WORKERS: number = 4;
    private idleTimeout: number = 10000; // 10s default
    private timeoutId: NodeJS.Timeout | null = null;

    private constructor() { }

    /**
     * Returns the singleton instance of the manager.
     */
    public static getInstance(): OcrSchedulerManager {
        if (!OcrSchedulerManager.instance) {
            OcrSchedulerManager.instance = new OcrSchedulerManager();
        }
        return OcrSchedulerManager.instance;
    }

    /**
     * Checks if the singleton instance has been initialized.
     */
    public static hasInstance(): boolean {
        return !!OcrSchedulerManager.instance;
    }

    /**
     * Resets the inactivity timer. If the timer reaches its duration, 
     * all workers are terminated automatically.
     */
    private resetIdleTimer(): void {
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
        }

        if (this.idleTimeout > 0) {
            this.timeoutId = setTimeout(async () => {
                await this.terminate();
            }, this.idleTimeout);
        }
    }

    /**
     * Performs OCR on an image using the smart worker pool.
     * 
     * @param image - Image data (Buffer, string path, or Blob)
     * @param config - OCR configuration (language, custom paths)
     * @returns Recognized text
     */
    public async recognize(image: any, config?: OcrConfig): Promise<string> {
        return new Promise((resolve, reject) => {
            // Update idle timeout if provided
            if (config?.autoTerminateTimeout !== undefined) {
                this.idleTimeout = config.autoTerminateTimeout;
            }

            // Reset the inactivity timer every time a new job is requested
            this.resetIdleTimer();

            // Add job to queue and trigger processing
            this.queue.push({ image, config: config || {}, resolve, reject });
            this.processQueue();
        });
    }

    /**
     * Attempts to process the next job in the queue using an available worker.
     */
    private async processQueue(): Promise<void> {
        if (this.queue.length === 0) return;

        const nextJob = this.queue[0];
        const requestedLanguage = nextJob.config.language || 'eng';

        // 1. Find an idle worker with the EXACT language affinity
        let managed = this.pool.find(mw => !mw.isBusy && mw.language === requestedLanguage);

        // 2. If not found and we have room, create a new worker
        if (!managed && this.pool.length < this.MAX_WORKERS) {
            try {
                const { createWorker } = await import('tesseract.js');
                const options: any = { logger: () => { } };
                if (nextJob.config.workerPath) options.workerPath = nextJob.config.workerPath;
                if (nextJob.config.corePath) options.corePath = nextJob.config.corePath;
                if (nextJob.config.langPath) options.langPath = nextJob.config.langPath;

                const worker = await createWorker(requestedLanguage, 1, options);
                managed = {
                    worker,
                    language: requestedLanguage,
                    lastUsed: Date.now(),
                    isBusy: false
                };
                this.pool.push(managed);
            } catch (err) {
                const job = this.queue.shift();
                job?.reject(err);
                this.processQueue(); // Try next job
                return;
            }
        }

        // 3. If still not found and we are at capacity, find the LRU idle worker and re-initialize it
        if (!managed) {
            const idleWorkers = this.pool.filter(mw => !mw.isBusy);
            if (idleWorkers.length > 0) {
                // Find Least Recently Used idle worker
                managed = idleWorkers.reduce((prev, curr) => (prev.lastUsed < curr.lastUsed ? prev : curr));
                
                try {
                    // Smart Re-initialization (v5 API)
                    await managed.worker.reinitialize(requestedLanguage);
                    managed.language = requestedLanguage;
                } catch (err) {
                    // If reinitialization fails, we might need to recreate it, but for simplicity 
                    // we'll just fail this job and try another worker next time.
                    const job = this.queue.shift();
                    job?.reject(err);
                    this.processQueue();
                    return;
                }
            }
        }

        // 4. If we have a worker ready, execute the job
        if (managed) {
            const job = this.queue.shift();
            if (!job) return;

            managed.isBusy = true;
            managed.lastUsed = Date.now();

            try {
                const { data: { text } } = await managed.worker.recognize(job.image);
                job.resolve(text);
            } catch (err) {
                job.reject(err);
            } finally {
                managed.isBusy = false;
                managed.lastUsed = Date.now();
                // Check if there are more jobs waiting
                this.processQueue();
            }
        }
        // If no worker is available (all busy), the job stays in the queue 
        // and will be picked up when a worker finishes.
    }

    /**
     * Terminates all workers in the pool and resets the state.
     */
    public async terminate(): Promise<void> {
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
            this.timeoutId = null;
        }

        const workersToTerminate = this.pool.map(mw => mw.worker.terminate());
        await Promise.all(workersToTerminate);
        this.pool = [];
    }
}

/**
 * Performs Optical Character Recognition (OCR) on an image to extract text.
 * 
 * Uses Tesseract.js to recognize text in the provided image buffer.
 * This is useful for extracting text from screenshots, scanned documents,
 * charts with labels, or any image containing text.
 * 
 * This function uses a shared worker pool to minimize initialization overhead.
 * 
 * @param image - The image data as a Buffer, file path, or Blob
 * @param config - Optional configuration for language and custom worker paths
 * @returns A promise that resolves to the recognized text as a string
 * @throws {Error} If the image cannot be processed or Tesseract initialization fails
 * 
 * @example
 * ```typescript
 * // Extract text from an English image
 * const text = await performOcr(imageBuffer, { language: 'eng' });
 * ```
 * 
 * @see https://github.com/naptha/tesseract.js for supported languages and options
 */
export const performOcr = async (image: Buffer | string, config?: OcrConfig): Promise<string> => {
    // Prepare image data
    let inputImage: any = image;

    // In browser environment, convert Buffer to Blob for better compatibility
    // @ts-ignore
    if (typeof window !== 'undefined' && typeof Blob !== 'undefined' && Buffer.isBuffer(image)) {
        inputImage = new Blob([image as any], { type: 'image/bmp' });
    }

    return await OcrSchedulerManager.getInstance().recognize(inputImage, config);
};

/**
 * Terminates all OCR workers and cleans up resources.
 * 
 * Should be called when the application is shutting down or OCR is no longer needed
 * to prevent memory leaks and dangling worker processes.
 */
export const terminateOcr = async (): Promise<void> => {
    if (OcrSchedulerManager.hasInstance()) {
        await OcrSchedulerManager.getInstance().terminate();
    }
};
