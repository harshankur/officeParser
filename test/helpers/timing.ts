/**
 * A case whose property is a time limit itself (pdf.js's maxTimeMs kept, an abort heard, work stopped as
 * soon as it can be, a long input read quickly): held to its limit, but run once more when over it and
 * the faster try kept, so a pause of the machine is not taken for the limit failing. A shared CI runner
 * pauses for seconds at a time, and on one Node version of a matrix at a time. The result is the last
 * try's, so `run` starts from nothing each time.
 *
 * Every limit on wall-clock time in the test suites goes through this, or through a suite's linear-time
 * helper, which measures again in the same way; never `Date.now() - started < limit` on one measurement.
 */
export const withinTime = async <T>(limitMs: number, run: () => T | Promise<T>): Promise<{ result: T; ms: number; within: boolean }> => {
    let started = performance.now();
    let result = await run();
    let ms = performance.now() - started;
    if (ms >= limitMs) {
        started = performance.now();
        result = await run();
        ms = Math.min(ms, performance.now() - started);
    }
    return { result, ms: Math.round(ms), within: ms < limitMs };
};
