import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const distDir = path.join(rootDir, 'dist');
const binDir = path.join(distDir, 'bin');

// Ensure directories exist
fs.mkdirSync(binDir, { recursive: true });

// Step 1: Bundle with esbuild
console.log('Bundling CLI with esbuild...');
execSync('npx esbuild src/cli.ts --bundle --platform=node --format=cjs --outfile=dist/cli-bundled.cjs --external:puppeteer --external:canvas', { cwd: rootDir, stdio: 'inherit' });

// Step 2: Write sea-config.json
console.log('Writing sea-config.json...');
const seaConfig = {
    main: "dist/cli-bundled.cjs",
    output: "dist/prep.blob"
};
fs.writeFileSync(path.join(distDir, 'sea-config.json'), JSON.stringify(seaConfig, null, 2), 'utf8');

// Step 3: Generate the blob
console.log('Generating SEA prep.blob...');
execSync('node --experimental-sea-config dist/sea-config.json', { cwd: rootDir, stdio: 'inherit' });

const postjectPath = path.join(rootDir, 'node_modules', '.bin', 'postject');

// Helper to inject blob
function injectBlob(targetPath, isMacho = false) {
    try {
        fs.chmodSync(targetPath, 0o755);
    } catch (e) {
        console.warn(`Warning: Could not set permissions on ${path.basename(targetPath)}`);
    }
    console.log(`Injecting SEA blob into ${path.basename(targetPath)}...`);
    const machoFlag = isMacho ? ' --macho-segment-name NODE_SEA' : '';
    execSync(`"${postjectPath}" "${targetPath}" NODE_SEA_BLOB dist/prep.blob --sentinel-fuse NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2${machoFlag}`, { cwd: rootDir, stdio: 'inherit' });
}

// Helper to download a file
async function downloadFile(url, dest) {
    if (fs.existsSync(dest)) {
        try { fs.chmodSync(dest, 0o755); } catch (e) {}
        fs.rmSync(dest, { force: true });
    }
    const res = await fetch(url);
    if (!res.ok) throw new Error(`Failed to download ${url}: ${res.statusText}`);
    const fileStream = fs.createWriteStream(dest);
    const body = res.body;
    const reader = body.getReader();
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        fileStream.write(value);
    }
    fileStream.end();
    return new Promise((resolve, reject) => {
        fileStream.on('finish', resolve);
        fileStream.on('error', reject);
    });
}

// Step 5: URLs and paths for target builds
const macosUrl = 'https://nodejs.org/dist/v25.9.0/node-v25.9.0-darwin-arm64.tar.gz';
const winUrl = 'https://nodejs.org/dist/v25.9.0/win-x64/node.exe';
const linuxUrl = 'https://nodejs.org/dist/v25.9.0/node-v25.9.0-linux-x64.tar.xz';

const macosTarPath = path.join(distDir, 'node-macos.tar.gz');
const macosArm64Dest = path.join(binDir, 'officeparser-macos-arm64');
const winDest = path.join(binDir, 'officeparser-win-x64.exe');
const linuxTarPath = path.join(distDir, 'node-linux.tar.xz');
const linuxDestNode = path.join(binDir, 'officeparser-linux-x64');

async function buildAll() {
    try {
        // Step 4: Build macOS ARM64 Executable
        console.log('Downloading macOS ARM64 Node package...');
        await downloadFile(macosUrl, macosTarPath);

        console.log('Extracting macOS ARM64 package...');
        execSync(`tar -xf "${macosTarPath}" -C "${distDir}"`, { stdio: 'inherit' });

        const macosSrcNode = path.join(distDir, 'node-v25.9.0-darwin-arm64', 'bin', 'node');
        if (fs.existsSync(macosArm64Dest)) {
            try { fs.chmodSync(macosArm64Dest, 0o755); } catch (e) {}
            fs.rmSync(macosArm64Dest, { force: true });
        }
        fs.copyFileSync(macosSrcNode, macosArm64Dest);

        console.log('Cleaning up temporary macOS files...');
        fs.rmSync(macosTarPath, { force: true });
        fs.rmSync(path.join(distDir, 'node-v25.9.0-darwin-arm64'), { recursive: true, force: true });

        console.log('Removing macOS code signature...');
        execSync(`codesign --remove-signature "${macosArm64Dest}"`, { stdio: 'inherit' });
        
        injectBlob(macosArm64Dest, true);
        
        console.log('Signing macOS binary...');
        execSync(`codesign --sign - "${macosArm64Dest}"`, { stdio: 'inherit' });

        // Step 5: Build Windows x64 Executable
        console.log('Downloading Windows x64 Node binary...');
        await downloadFile(winUrl, winDest);
        injectBlob(winDest);

        // Step 6: Build Linux x64 Executable
        console.log('Downloading Linux x64 Node package...');
        await downloadFile(linuxUrl, linuxTarPath);

        console.log('Extracting Linux x64 package...');
        execSync(`tar -xf "${linuxTarPath}" -C "${distDir}"`, { stdio: 'inherit' });

        const linuxSrcNode = path.join(distDir, 'node-v25.9.0-linux-x64', 'bin', 'node');
        if (fs.existsSync(linuxDestNode)) {
            try { fs.chmodSync(linuxDestNode, 0o755); } catch (e) {}
            fs.rmSync(linuxDestNode, { force: true });
        }
        fs.copyFileSync(linuxSrcNode, linuxDestNode);

        injectBlob(linuxDestNode);

        // Clean up temp linux extraction
        console.log('Cleaning up temporary Linux files...');
        fs.rmSync(linuxTarPath, { force: true });
        fs.rmSync(path.join(distDir, 'node-v25.9.0-linux-x64'), { recursive: true, force: true });

        console.log('🎉 Successfully created all binaries in dist/bin/ !');
    } catch (err) {
        console.error('Error building binaries:', err);
        process.exit(1);
    }
}

buildAll();
