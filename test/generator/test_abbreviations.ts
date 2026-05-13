import { OfficeParser } from '../../src/OfficeParser';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParserAST } from '../../src/types';

async function testAbbreviations() {
    console.log('Testing custom abbreviations...');
    
    // Mock AST with text containing an abbreviation
    const ast: any = {
        type: 'docx',
        content: [
            {
                type: 'paragraph',
                text: 'The company name is Alpha Inc. It is located in the city.'
            }
        ],
        toText: () => 'The company name is Alpha Inc. It is located in the city.'
    };

    // Test 1: Default (should NOT split on Inc.)
    const resultDefault = await OfficeGenerator.generate(ast as any, 'chunks', {
        chunksConfig: { strategy: 'document-structure' }
    });
    const chunksDefault = resultDefault.value as any[];
    console.log('Default chunks:', chunksDefault.map(c => c.text));

    // Test 2: Custom (add "Alpha" as abbreviation)
    const resultCustom = await OfficeGenerator.generate(ast as any, 'chunks', {
        chunksConfig: { 
            strategy: 'document-structure',
            abbreviations: ['Mr', 'Dr', 'Ms', 'Inc', 'Ltd', 'Prof', 'Sr', 'Jr', 'vs', 'etc', 'Alpha']
        }
    });
    const chunksCustom = resultCustom.value as any[];
    console.log('Custom chunks (with Alpha):', chunksCustom.map(c => c.text));
    
    // Verify that Inc. didn't split in both cases because it's in the default list
    if (chunksDefault.length === 1 && chunksDefault[0].text.includes('Alpha Inc. It')) {
        console.log('✓ PASS: Default abbreviations worked.');
    } else {
        console.error('✗ FAIL: Default abbreviations failed.');
    }
}

testAbbreviations().catch(console.error);
