/**
 * Automated License Compliance Validator
 * 
 * This script scans the generated sbom.cdx.json file for restricted "Copyleft"
 * or "Source-Available" licenses that are incompatible with this project's 
 * MIT license.
 * 
 * It will fail the build if any violations are found.
 */

const fs = require('fs');
const path = require('path');

const SBOM_PATH = path.join(__dirname, '..', 'dist', 'sbom.cdx.json');

// List of restricted license patterns (Regex-based IDs)
const RESTRICTED_PATTERNS = [
    /GPL/i,             // General Public License (v2, v3, etc.)
    /LGPL/i,            // Lesser General Public License
    /AGPL/i,            // Affero General Public License
    /SSPL/i,            // Server Side Public License
    /BSL-1\.1/i,        // Business Source License
    /CC-BY-NC/i,        // Creative Commons Non-Commercial
    /JSON/i,            // The "Good not Evil" license
    /WTFPL/i,           // Risk-heavy (legal disclaimer missing)
    /Commons-Clause/i   // Commons Clause restricted projects
];

/**
 * Checks if a license ID or name matches a restricted pattern.
 */
function isRestricted(license) {
    if (!license) return false;
    const id = license.id || '';
    const name = license.name || '';
    return RESTRICTED_PATTERNS.some(p => p.test(id) || p.test(name));
}

async function main() {
    console.log(`\n🔍 Validating License Compliance from SBOM: ${SBOM_PATH}\n`);

    if (!fs.existsSync(SBOM_PATH)) {
        console.error(`❌ SBOM file not found at ${SBOM_PATH}. Did you run 'npm run sbom' first?`);
        process.exit(1);
    }

    let bom;
    try {
        bom = JSON.parse(fs.readFileSync(SBOM_PATH, 'utf8'));
    } catch (e) {
        console.error(`❌ Failed to parse SBOM: ${e.message}`);
        process.exit(1);
    }

    const violations = [];
    const components = bom.components || [];

    for (const comp of components) {
        // Skip Dev dependencies as they are not bundled into production artifacts
        const isDev = comp.properties?.some(p => p.name === 'cdx:npm:package:development' && p.value === 'true');
        if (isDev) continue;

        const licenses = comp.licenses || [];
        for (const lEntry of licenses) {
            const license = lEntry.license;
            if (isRestricted(license)) {
                violations.push({
                    name: comp.name,
                    version: comp.version,
                    license: license.id || license.name || 'Unknown'
                });
            }
        }
    }

    if (violations.length > 0) {
        console.error('═'.repeat(70));
        console.error('  ❌ LICENSE VIOLATIONS DETECTED');
        console.error('═'.repeat(70));
        console.error('  The following PRODUCTION dependencies use restricted licenses:');

        violations.forEach(v => {
            console.error(`  - ${v.name}@${v.version} (License: ${v.license})`);
        });

        console.error('\n  Action Required: Replace these dependencies with MIT/Apache compatible alternatives.');
        process.exit(1);
    }

    console.log('✅ All production dependencies have compliant licenses.');
    process.exit(0);
}

main().catch(err => {
    console.error('❌ Validator error:', err);
    process.exit(1);
});
