/**
 * Browser-side stub for 'puppeteer'.
 * 
 * PDF generation in the browser uses native window.print() 
 * instead of Puppeteer.
 */

export const launch = () => {
    throw new Error("officeparser: 'puppeteer.launch' is not supported in the browser. Browser-side PDF generation uses native print capabilities.");
};

export default {
    launch
};
