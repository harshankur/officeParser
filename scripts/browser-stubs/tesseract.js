/**
 * Browser-side stub for 'tesseract.js' in the slim browser bundle.
 * 
 * OCR is disabled/unsupported in the slim bundle.
 */

export const createWorker = () => {
    throw new Error("officeparser: OCR (tesseract.js) is disabled in the slim browser bundle. Please use the full browser bundle if you need OCR support.");
};

export default {
    createWorker
};
