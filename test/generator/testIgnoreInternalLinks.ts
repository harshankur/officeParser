import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParserAST } from '../../src/types';
import * as assert from 'assert';

async function runTests() {
    console.log('Running ignoreInternalLinks tests...');

    const mockAST: OfficeParserAST = {
        type: 'docx',
        metadata: { title: 'Test Doc' },
        attachments: [],
        content: [
            {
                type: 'heading',
                text: 'Main Heading',
                metadata: { level: 1, anchorIds: ['manual-id'] } as any,
                children: [
                    {
                        type: 'text',
                        text: 'Main Heading',
                    },
                    {
                        type: 'text',
                        text: ' with ',
                    },
                    {
                        type: 'text',
                        text: 'Internal Link',
                        metadata: { link: '#manual-id', linkType: 'internal' } as any
                    },
                    {
                        type: 'text',
                        text: ' and ',
                    },
                    {
                        type: 'text',
                        text: 'External Link',
                        metadata: { link: 'https://google.com', linkType: 'external' } as any
                    }
                ]
            }
        ],
        toText: () => 'Main Heading with Internal Link and External Link',
        getImages: () => []
    } as any;

    // Test HTML
    console.log('- Testing HTML...');
    const htmlDefault = (await OfficeGenerator.generate(mockAST, 'html')).value as string;
    assert.ok(htmlDefault.includes('id="manual-id"'), 'Default HTML should include manual-id');
    assert.ok(htmlDefault.includes('href="#manual-id"'), 'Default HTML should include internal link');
    assert.ok(htmlDefault.includes('href="https://google.com"'), 'Default HTML should include external link');

    const htmlIgnore = (await OfficeGenerator.generate(mockAST, 'html', { ignoreInternalLinks: true })).value as string;
    assert.ok(!htmlIgnore.includes('id="manual-id"'), 'IgnoreInternalLinks HTML should NOT include manual-id');
    assert.ok(!htmlIgnore.includes('href="#manual-id"'), 'IgnoreInternalLinks HTML should NOT include internal link');
    assert.ok(htmlIgnore.includes('href="https://google.com"'), 'IgnoreInternalLinks HTML should STILL include external link');

    // Test Markdown
    console.log('- Testing Markdown...');
    const mdDefault = (await OfficeGenerator.generate(mockAST, 'md')).value as string;
    assert.ok(mdDefault.includes('{#manual-id}'), 'Default MD should include {#manual-id}');
    assert.ok(mdDefault.includes('[Internal Link](#manual-id)'), 'Default MD should include internal link');
    assert.ok(mdDefault.includes('[External Link](https://google.com)'), 'Default MD should include external link');

    const mdIgnore = (await OfficeGenerator.generate(mockAST, 'md', { ignoreInternalLinks: true })).value as string;
    assert.ok(!mdIgnore.includes('{#manual-id}'), 'IgnoreInternalLinks MD should NOT include {#manual-id}');
    assert.ok(mdIgnore.includes('Main Heading'), 'IgnoreInternalLinks MD should still have heading text');
    assert.ok(!mdIgnore.includes('[Internal Link](#manual-id)'), 'IgnoreInternalLinks MD should NOT include internal link');
    assert.ok(mdIgnore.includes('Internal Link'), 'IgnoreInternalLinks MD should still have link text');
    assert.ok(mdIgnore.includes('[External Link](https://google.com)'), 'IgnoreInternalLinks MD should STILL include external link');

    // Test RTF
    console.log('- Testing RTF...');
    const rtfDefault = (await OfficeGenerator.generate(mockAST, 'rtf')).value as string;
    assert.ok(rtfDefault.includes('HYPERLINK "#manual-id"'), 'Default RTF should include internal link');
    assert.ok(rtfDefault.includes('HYPERLINK "https://google.com"'), 'Default RTF should include external link');

    const rtfIgnore = (await OfficeGenerator.generate(mockAST, 'rtf', { ignoreInternalLinks: true })).value as string;
    assert.ok(!rtfIgnore.includes('HYPERLINK "#manual-id"'), 'IgnoreInternalLinks RTF should NOT include internal link');
    assert.ok(rtfIgnore.includes('HYPERLINK "https://google.com"'), 'IgnoreInternalLinks RTF should STILL include external link');

    console.log('All ignoreInternalLinks tests passed!');
}

runTests().catch(err => {
    console.error('Test failed:', err);
    process.exit(1);
});
