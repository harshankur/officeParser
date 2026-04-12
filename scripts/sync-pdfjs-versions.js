const fs = require('fs');
const path = require('path');

/**
 * Synchronization Script for PDF.js Version
 * 
 * This script ensures that all documentation, examples, and type definitions
 * use the same version of pdfjs-dist as specified in package.json.
 */

function syncVersions() {
    const args = process.argv.slice(2);
    const versionFromArg = args.find(arg => !arg.startsWith('--'));

    const packageJsonPath = path.join(__dirname, '../package.json');
    if (!fs.existsSync(packageJsonPath)) {
        console.error('Error: package.json not found');
        process.exit(1);
    }

    let packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

    // If version is provided as an argument, update package.json first
    if (versionFromArg) {
        console.log(`Updating package.json pdfjs-dist dependency to: ${versionFromArg}`);
        packageJson.dependencies['pdfjs-dist'] = versionFromArg;
        fs.writeFileSync(packageJsonPath, JSON.stringify(packageJson, null, 2) + '\n', 'utf8');
    }

    const pdfJsDependency = packageJson.dependencies['pdfjs-dist'];

    if (!pdfJsDependency) {
        console.error('Error: pdfjs-dist not found in dependencies');
        process.exit(1);
    }

    // Strip characters like ^, ~ from version
    const pdfJsVersion = pdfJsDependency.replace(/[\^~]/g, '');
    const versionRegex = /pdfjs-dist@[\d.]+/g;
    const targetVersionString = `pdfjs-dist@${pdfJsVersion}`;

    console.log(`Syncing PDF.js version: ${pdfJsVersion}`);

    const filesToUpdate = [
        'src/types.ts',
        'docs/index.html',
        'docs/visualizer_old.html',
        'README.md'
    ];

    let updatedCount = 0;
    const handledFiles = new Set(filesToUpdate.map(f => path.normalize(f)));

    // 1. Perform Sync
    filesToUpdate.forEach(relativePath => {
        const filePath = path.join(__dirname, '..', relativePath);
        if (fs.existsSync(filePath)) {
            const content = fs.readFileSync(filePath, 'utf8');
            const newContent = content.replace(versionRegex, targetVersionString);

            if (content !== newContent) {
                fs.writeFileSync(filePath, newContent, 'utf8');
                console.log(`  ✓ Updated ${relativePath}`);
                updatedCount++;
            } else {
                console.log(`  - ${relativePath} is already up to date`);
            }
        } else {
            console.warn(`  ! Warning: File not found: ${relativePath}`);
        }
    });

    // 2. Scan for unhandled occurrences
    console.log('\nScanning for unhandled PDF.js version occurrences...');
    const rootDir = path.join(__dirname, '..');
    const ignoreDirs = ['node_modules', 'dist', '.git', 'test/results'];
    const unhandledFiles = [];

    function scanDir(currentDir) {
        const files = fs.readdirSync(currentDir);
        for (const file of files) {
            const fullPath = path.join(currentDir, file);
            const relPath = path.relative(rootDir, fullPath);

            if (fs.statSync(fullPath).isDirectory()) {
                if (!ignoreDirs.includes(file)) {
                    scanDir(fullPath);
                }
                continue;
            }

            // Only scan text-like files
            if (!/\.(ts|js|html|md|json|txt|css)$/.test(file)) continue;

            // Skip the package files and the script itself
            if (relPath === 'package.json' || relPath === 'package-lock.json') continue;
            if (relPath === 'scripts/sync-pdfjs-versions.js') continue;

            try {
                const content = fs.readFileSync(fullPath, 'utf8');
                // Use a fresh regex for test to avoid lastIndex issues
                if (/pdfjs-dist@[\d.]+/.test(content) && !handledFiles.has(path.normalize(relPath))) {
                    unhandledFiles.push(relPath);
                }
            } catch (err) {
                console.warn(`  ! Warning: Could not read ${relPath}: ${err.message}`);
            }
        }
    }

    try {
        scanDir(rootDir);
    } catch (err) {
        console.error(`Error during scan: ${err.message}`);
    }

    if (unhandledFiles.length > 0) {
        console.warn('\n[!] WARNING: Found PDF.js version strings in files not handled by this script:');
        unhandledFiles.forEach(f => console.warn(`    - ${f}`));
        console.warn('Please add these files to the "filesToUpdate" list in scripts/sync-pdfjs-versions.js if they need synchronization.\n');
    } else {
        console.log('  ✓ No unhandled occurrences found.');
    }

    console.log(`\nSync complete. ${updatedCount} file(s) updated.`);
}

syncVersions();
