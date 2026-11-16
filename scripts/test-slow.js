#!/usr/bin/env node
/**
 * The slow runner: the full test suite in a Linux container whose CPU is held to about a GitHub-hosted
 * runner's speed, so a limit that passes only on a fast machine fails here, before a release, rather than
 * in the release's own workflow.
 *
 *   npm run test:slow                      the full suite (npm test) on the working tree as git sees it
 *   npm run test:slow -- --committed       on HEAD only, without uncommitted changes
 *   npm run test:slow -- --script test:fast
 *   npm run test:slow -- --cpus 4 --memory 6g --platform linux/arm64
 *
 * Needs Docker. The image (test/slow-runner/Dockerfile) holds only the dependencies and is built again
 * when package-lock.json or the Dockerfile changes. The files are copied into the container (tracked and
 * untracked files git does not ignore, or HEAD's with --committed), so a run writes nothing into the
 * repository, and the container is removed when it ends or is stopped.
 *
 * The container is x86 (linux/amd64), as GitHub's runners are, which Docker Desktop on an Apple Silicon
 * Mac runs through Rosetta: every core is slower, about as GitHub's are, and Node's garbage collector
 * still has cores to work on in parallel. (A CPU cap on a native ARM container was tried first: at 1.1
 * CPUs the chart write below took GitHub's time, but its threads shared one core, so the collector
 * worked as on one core and three growth checks that pass on GitHub failed.) CPUS is 2, as a GitHub
 * runner's 4 vCPUs are 2 cores with two threads each. Measured with the Mac otherwise idle (November
 * 2026): the 1,000 copies of a chart written to native PDF in test/security/chartWritersHardening.ts took
 * 17.6 s on an Apple M1 Pro, 49 to 50 s on GitHub (Node 24), and 51 s here (41 s at 4 CPUs). Measure
 * again when the image or GitHub's runners change. On a machine with no Rosetta, Docker emulates x86 far
 * more slowly: there, `--platform linux/arm64 --cpus 1.1` is the nearer match.
 */
const { execFileSync, spawn } = require('child_process');
const crypto = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DOCKERFILE = path.join(ROOT, 'test', 'slow-runner', 'Dockerfile');
const CPUS = '2';
const PLATFORM = 'linux/amd64';
const MEMORY = '7g';

const args = process.argv.slice(2);
const option = (name, fallback) => {
    const i = args.indexOf(name);
    if (i < 0) return fallback;
    if (i + 1 >= args.length) throw new Error(`${name} needs a value`);
    return args[i + 1];
};
const known = new Set(['--committed', '--script', '--cpus', '--memory', '--platform']);
for (let i = 0; i < args.length; i++) {
    if (!known.has(args[i])) { console.error(`Unknown option ${args[i]} (known: ${[...known].join(', ')})`); process.exit(2); }
    if (args[i] !== '--committed') i++;
}
const committed = args.includes('--committed');
const script = option('--script', 'test');
const cpus = option('--cpus', CPUS);
const memory = option('--memory', MEMORY);
const platform = option('--platform', PLATFORM);
if (!/^[\w:.-]+$/.test(script)) { console.error(`--script must be an npm script name, not "${script}"`); process.exit(2); }

const run = (cmd, cmdArgs, opts = {}) => execFileSync(cmd, cmdArgs, { cwd: ROOT, encoding: 'utf8', maxBuffer: 256 * 1024 * 1024, ...opts });

try { run('docker', ['version', '--format', '{{.Server.Version}}'], { stdio: 'pipe' }); }
catch { console.error('The slow runner needs Docker, with its daemon running.'); process.exit(1); }

// The image, named by what it is built from.
const hash = crypto.createHash('sha256');
hash.update(platform);
for (const file of [DOCKERFILE, path.join(ROOT, 'package.json'), path.join(ROOT, 'package-lock.json')]) hash.update(fs.readFileSync(file));
const image = `officeparser-slow-runner:${hash.digest('hex').slice(0, 12)}`;
let built = true;
try { run('docker', ['image', 'inspect', image], { stdio: 'pipe' }); } catch { built = false; }
if (!built) {
    console.log(`Building ${image} for ${platform} (the dependencies, once for each package-lock.json)...`);
    const context = fs.mkdtempSync(path.join(os.tmpdir(), 'officeparser-slow-runner-'));
    try {
        for (const file of ['package.json', 'package-lock.json']) fs.copyFileSync(path.join(ROOT, file), path.join(context, file));
        fs.copyFileSync(DOCKERFILE, path.join(context, 'Dockerfile'));
        run('docker', ['build', '--platform', platform, '--tag', image, context], { stdio: 'inherit' });
    } finally {
        fs.rmSync(context, { recursive: true, force: true });
    }
}

// The files to test, as one tar on the container's standard input.
let tar;
if (committed) {
    tar = Buffer.from(run('git', ['archive', '--format=tar', 'HEAD'], { encoding: 'buffer' }));
} else {
    const listed = run('git', ['ls-files', '-z', '--cached', '--others', '--exclude-standard', '--deduplicate']).split('\0').filter(Boolean);
    // A tracked file deleted in the working tree is not copied, as it is not there.
    const present = listed.filter(file => fs.existsSync(path.join(ROOT, file)));
    const dirty = run('git', ['status', '--porcelain']).trim();
    if (dirty) console.log('Testing the working tree, with its uncommitted changes (--committed tests HEAD only).');
    // macOS's tar would add its extended attributes, which Linux's tar warns of for every file.
    const quiet = process.platform === 'darwin' ? ['--no-mac-metadata', '--no-xattrs'] : [];
    tar = Buffer.from(run('tar', ['-c', '-f', '-', ...quiet, '--null', '-T', '-'], { input: Buffer.from(present.join('\0')), encoding: 'buffer', env: { ...process.env, COPYFILE_DISABLE: '1' } }));
}

const name = `officeparser-slow-runner-${process.pid}`;
console.log(`Running npm run ${script} in ${image} (${platform}, --cpus ${cpus}, --memory ${memory})...`);
const started = Date.now();
// Node sizes its heap from the memory it sees: 2 GB in a 7 GB container, where a GitHub runner's 16 GB
// (and a Mac's) give it 4 GB, and the suites need it. So the heap is a GitHub runner's.
const child = spawn('docker', ['run', '--rm', '-i', '--init', '--name', name, '--platform', platform, '--cpus', cpus, '--memory', memory, '--memory-swap', memory,
    '--env', 'NODE_OPTIONS=--max-old-space-size=4096',
    image, 'sh', '-c', `tar -x -f - -C /work && npm run ${script}`], { cwd: ROOT, stdio: ['pipe', 'inherit', 'inherit'] });
child.stdin.end(tar);

// Stopped here: the container is stopped too, and removed by --rm.
let stopping = false;
const stop = () => {
    if (stopping) return;
    stopping = true;
    try { execFileSync('docker', ['stop', '--time', '5', name], { stdio: 'ignore' }); } catch { /* already gone */ }
};
process.on('SIGINT', stop);
process.on('SIGTERM', stop);

child.on('exit', (code, signal) => {
    const minutes = ((Date.now() - started) / 60000).toFixed(1);
    console.log(`Slow runner: npm run ${script} ${code === 0 ? 'passed' : 'failed'} in ${minutes} minutes${signal ? ` (${signal})` : ''}.`);
    process.exit(code ?? 1);
});
