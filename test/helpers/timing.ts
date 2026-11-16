/**
 * Wall-clock limits in the test suites, scaled to the machine running them.
 *
 * A limit written as a number of milliseconds measures the machine as much as the code: a check that takes
 * 18 seconds on the laptop the limits were written on took 50 on a GitHub runner. So each test process
 * first times a fixed piece of the library's own work (calibrationMs) and compares it with the time it
 * takes on that laptop (REFERENCE_MS): a machine that takes three times as long gets limits three times as
 * long. The factor is never below 1, so a faster machine keeps the written limits rather than tighter ones,
 * and it is printed once, so a failure reads as "limit 20000 ms x 2.8".
 *
 * Three kinds of time limit, and which of them scale:
 * - growth (n against 4n, 250 copies against 1,000): two measurements on the same machine, so no factor,
 *   except an absolute slack added to the ratio (`8 * small + 250`), which is time and scales (`scaledMs`);
 * - "in reasonable time" (a hostile file read in seconds): `withinTime`, the whole limit scaled;
 * - a deadline the code is given and must keep (`maxTimeMs`, an abort signal): `withinDeadline`, where the
 *   deadline is the property and stays as it is, and only the slack for the work around it scales.
 *
 * Every limit on wall-clock time in the test suites goes through these, or through a suite's growth helper,
 * which measures again in the same way; never `Date.now() - started < limit` on one measurement.
 *
 * OFFICEPARSER_TEST_SPEED_FACTOR sets the factor instead of measuring it (a number above 0; below 1 only
 * when set this way), for a machine whose calibration is not to be trusted or to try a suite's limits.
 */
import * as fs from 'fs';
import * as path from 'path';

/**
 * What calibrationMs takes on the machine the suites' limits were written on (an Apple M1 Pro, Node 24):
 * the median of its timed rounds, measured in fresh processes.
 */
export const REFERENCE_MS = 125;

/** The factor is held to this at most, so a calibration caught in a long pause does not turn every limit off. */
const MAX_FACTOR = 20;

/**
 * One round of the calibration: a Word document written to native PDF, a deck and a workbook written to
 * HTML, from bytes already in memory. It runs the zip, XML, layout and writer code most limits time.
 */
const calibrationRound = async (): Promise<void> => {
    const { OfficeParser } = await import('../../src/OfficeParser');
    const quiet = { onWarning: () => { } } as any;
    for (const [name, format, config] of [['test.docx', 'pdf', { pdfConfig: { engine: 'native' } }], ['test.pptx', 'html', {}], ['test.xlsx', 'html', {}]] as const) {
        const ast = await OfficeParser.parseOffice(calibrationFiles()[name], quiet) as any;
        await ast.to(format, { ...quiet, ...config });
    }
};

let files: Record<string, Buffer> | undefined;
const calibrationFiles = (): Record<string, Buffer> => files ??= Object.fromEntries(
    ['test.docx', 'test.pptx', 'test.xlsx'].map(name => [name, fs.readFileSync(path.join(__dirname, '..', 'files', name))]));

/** The median time of the calibration's timed rounds, after two untimed rounds that warm the code up. */
export const calibrationMs = async (): Promise<number> => {
    calibrationFiles();
    for (let i = 0; i < 2; i++) await calibrationRound();
    const times: number[] = [];
    for (let i = 0; i < 5; i++) {
        const started = performance.now();
        await calibrationRound();
        times.push(performance.now() - started);
    }
    times.sort((a, b) => a - b);
    return times[2];
};

let factor: Promise<number> | undefined;

/** How many times longer than the reference machine this one takes: measured once a process, printed once. */
export const speedFactor = (): Promise<number> => factor ??= (async () => {
    const set = process.env.OFFICEPARSER_TEST_SPEED_FACTOR;
    if (set !== undefined && set !== '') {
        const value = Number(set);
        if (!(value > 0) || !Number.isFinite(value)) throw new Error(`OFFICEPARSER_TEST_SPEED_FACTOR must be a number above 0, not "${set}"`);
        console.log(`  (timing: limits x ${value}, set by OFFICEPARSER_TEST_SPEED_FACTOR)`);
        return value;
    }
    const ms = await calibrationMs();
    const value = Math.min(MAX_FACTOR, Math.max(1, ms / REFERENCE_MS));
    console.log(`  (timing: calibration ${Math.round(ms)} ms against ${REFERENCE_MS} ms on the reference machine: limits x ${value.toFixed(2)})`);
    return value;
})();

/** `ms` of the reference machine's time, on this machine. */
export const scaledMs = async (ms: number): Promise<number> => ms * await speedFactor();

type Timed<T> = { result: T; ms: number; within: boolean; limitMs: number; detail: string };

/** `run` timed, and run once more when over `limitMs` with the faster try kept (the result is the last try's). */
const timedAgainst = async <T>(limitMs: number, explain: string, run: () => T | Promise<T>): Promise<Timed<T>> => {
    let started = performance.now();
    let result = await run();
    let ms = performance.now() - started;
    if (ms >= limitMs) {
        // A shared CI runner pauses for seconds at a time, on one Node version of a matrix at a time: a pause
        // of the machine is not taken for the limit failing. `run` starts from nothing each time.
        started = performance.now();
        result = await run();
        ms = Math.min(ms, performance.now() - started);
    }
    return { result, ms: Math.round(ms), within: ms < limitMs, limitMs: Math.round(limitMs), detail: `${Math.round(ms)} ms, limit ${Math.round(limitMs)} ms (${explain})` };
};

/**
 * A case whose property is that its work is done in reasonable time (a hostile file read in seconds, work
 * stopped as soon as it can be): `limitMs` of the reference machine's time, scaled to this one.
 */
export const withinTime = async <T>(limitMs: number, run: () => T | Promise<T>): Promise<Timed<T>> => {
    const f = await speedFactor();
    return timedAgainst(limitMs * f, `${limitMs} ms x ${f.toFixed(2)}`, run);
};

/**
 * A case whose property is a deadline the code is given and keeps (pdf.js's maxTimeMs, an abort heard):
 * the deadline is wall-clock time on any machine, so only `slackMs`, the time for the work around it
 * (starting the parse, unwinding after the stop), is scaled.
 */
export const withinDeadline = async <T>(deadlineMs: number, slackMs: number, run: () => T | Promise<T>): Promise<Timed<T>> => {
    const f = await speedFactor();
    return timedAgainst(deadlineMs + slackMs * f, `deadline ${deadlineMs} ms + ${slackMs} ms x ${f.toFixed(2)}`, run);
};
