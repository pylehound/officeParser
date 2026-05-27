/**
 * Unit tests for Word comment and tracked-change extraction.
 *
 * Creates synthetic minimal .docx buffers in memory so the tests have
 * no external file dependency and are fully self-contained.
 */

import * as assert from 'assert';
// @ts-ignore - jszip is a devDependency installed for tests only
import JSZip from 'jszip';
import { OfficeParser } from '../src/OfficeParser';
import { CommentBlock, DeletionBlock, InsertionBlock } from '../src/types';

// ---------------------------------------------------------------------------
// DOCX fixture builder
// ---------------------------------------------------------------------------

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const WORD_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

async function buildDocx(documentXml: string, commentsXml: string): Promise<Buffer> {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CONTENT_TYPES);
    zip.file('_rels/.rels', RELS);
    zip.file('word/_rels/document.xml.rels', WORD_RELS);
    zip.file('word/document.xml', documentXml);
    zip.file('word/comments.xml', commentsXml);
    const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    return buffer;
}

// ---------------------------------------------------------------------------
// Fixture: document with one comment (no reply)
// ---------------------------------------------------------------------------

const COMMENT_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Normal text. </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>Anchored clause</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r>
        <w:commentReference w:id="1"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const COMMENT_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="1" w:author="Jane Doe" w:date="2026-01-15T10:00:00Z">
    <w:p w14:paraId="aabb0001">
      <w:r><w:t>Review this clause carefully.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`;

// ---------------------------------------------------------------------------
// Fixture: document with comment + reply
// ---------------------------------------------------------------------------

const REPLY_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="1" w:author="Jane Doe" w:date="2026-01-15T10:00:00Z">
    <w:p w14:paraId="aabb0001">
      <w:r><w:t>Review this clause carefully.</w:t></w:r>
    </w:p>
  </w:comment>
  <w:comment w:id="2" w:author="John Smith" w:date="2026-01-16T11:00:00Z">
    <w:p w14:paraId="aabb0002" w14:paraIdParent="aabb0001">
      <w:r><w:t>Agreed, narrowing to direct damages.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`;

// ---------------------------------------------------------------------------
// Fixture: document with tracked insertion and deletion
// ---------------------------------------------------------------------------

const TRACKED_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">The parties agree </w:t></w:r>
      <w:ins w:id="10" w:author="Jane Doe" w:date="2026-01-15T10:00:00Z">
        <w:r><w:t>unconditionally </w:t></w:r>
      </w:ins>
      <w:del w:id="11" w:author="John Smith" w:date="2026-01-16T11:00:00Z">
        <w:r><w:delText>to pay all costs</w:delText></w:r>
      </w:del>
      <w:r><w:t>to the terms.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

const EMPTY_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

// ---------------------------------------------------------------------------
// Fixture: comment anchored to text inside a table cell
// ---------------------------------------------------------------------------

const TABLE_COMMENT_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:commentRangeStart w:id="5"/>
            <w:r><w:t>Cell text</w:t></w:r>
            <w:commentRangeEnd w:id="5"/>
            <w:r><w:commentReference w:id="5"/></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

const TABLE_COMMENT_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="5" w:author="Alice" w:date="2026-01-20T10:00:00Z">
    <w:p w14:paraId="bbcc0005">
      <w:r><w:t>Verify this cell value.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`;

// ---------------------------------------------------------------------------
// Fixture: comment range spanning two paragraphs
// ---------------------------------------------------------------------------

const MULTI_PARA_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="6"/>
      <w:r><w:t>First paragraph text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">Second paragraph text.</w:t></w:r>
      <w:commentRangeEnd w:id="6"/>
      <w:r><w:commentReference w:id="6"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

const MULTI_PARA_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="6" w:author="Bob" w:date="2026-01-21T10:00:00Z">
    <w:p w14:paraId="bbcc0006">
      <w:r><w:t>Spans two paragraphs.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`;

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

async function runTests(): Promise<void> {
    let passed = 0;
    let failed = 0;

    async function test(name: string, fn: () => Promise<void>): Promise<void> {
        try {
            await fn();
            console.log(`  ✓ ${name}`);
            passed++;
        } catch (e: any) {
            console.error(`  ✗ ${name}`);
            console.error(`    ${e.message}`);
            failed++;
        }
    }

    console.log('\nWord Comments & Tracked Changes Tests\n');

    await test('extracts a comment block with author, date, and anchor text', async () => {
        const buf = await buildDocx(COMMENT_DOCUMENT_XML, COMMENT_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        const comments = (ast.blocks ?? []).filter(b => b.type === 'comment') as CommentBlock[];

        assert.strictEqual(comments.length, 1);
        assert.strictEqual(comments[0].text, 'Review this clause carefully.');
        assert.strictEqual(comments[0].author, 'Jane Doe');
        assert.ok(comments[0].date?.startsWith('2026-01-15'));
        assert.strictEqual(comments[0].anchorText, 'Anchored clause');
        assert.strictEqual(comments[0].replies?.length ?? 0, 0);
    });

    await test('attaches reply to parent comment', async () => {
        const buf = await buildDocx(COMMENT_DOCUMENT_XML, REPLY_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        const comments = (ast.blocks ?? []).filter(b => b.type === 'comment') as CommentBlock[];

        assert.strictEqual(comments.length, 1, 'only top-level comment emitted as block');
        const replies = comments[0].replies ?? [];
        assert.strictEqual(replies.length, 1);
        assert.strictEqual(replies[0].text, 'Agreed, narrowing to direct damages.');
        assert.strictEqual(replies[0].author, 'John Smith');
    });

    await test('suppresses comments when ignoreComments is true', async () => {
        const buf = await buildDocx(COMMENT_DOCUMENT_XML, COMMENT_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf, { ignoreComments: true });
        const comments = (ast.blocks ?? []).filter(b => b.type === 'comment');
        assert.strictEqual(comments.length, 0);
    });

    await test('extracts insertion block with author and date', async () => {
        const buf = await buildDocx(TRACKED_DOCUMENT_XML, EMPTY_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        const insertions = (ast.blocks ?? []).filter(b => b.type === 'insertion') as InsertionBlock[];

        assert.strictEqual(insertions.length, 1);
        assert.ok(insertions[0].text.includes('unconditionally'));
        assert.strictEqual(insertions[0].author, 'Jane Doe');
        assert.ok(insertions[0].date?.startsWith('2026-01-15'));
    });

    await test('extracts deletion block with author and date', async () => {
        const buf = await buildDocx(TRACKED_DOCUMENT_XML, EMPTY_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        const deletions = (ast.blocks ?? []).filter(b => b.type === 'deletion') as DeletionBlock[];

        assert.strictEqual(deletions.length, 1);
        assert.strictEqual(deletions[0].text, 'to pay all costs');
        assert.strictEqual(deletions[0].author, 'John Smith');
        assert.ok(deletions[0].date?.startsWith('2026-01-16'));
    });

    await test('suppresses tracked changes when ignoreTrackedChanges is true', async () => {
        const buf = await buildDocx(TRACKED_DOCUMENT_XML, EMPTY_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf, { ignoreTrackedChanges: true });
        const tracked = (ast.blocks ?? []).filter(b => b.type === 'insertion' || b.type === 'deletion');
        assert.strictEqual(tracked.length, 0);
    });

    await test('includes inserted text in fullText but not deleted text', async () => {
        const buf = await buildDocx(TRACKED_DOCUMENT_XML, EMPTY_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        assert.ok(ast.fullText?.includes('unconditionally'), 'inserted text in fullText');
        assert.ok(!ast.fullText?.includes('to pay all costs'), 'deleted text not in fullText');
    });

    await test('emits comment anchored inside a table cell', async () => {
        const buf = await buildDocx(TABLE_COMMENT_DOCUMENT_XML, TABLE_COMMENT_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        const comments = (ast.blocks ?? []).filter(b => b.type === 'comment') as CommentBlock[];

        assert.strictEqual(comments.length, 1, 'comment inside table cell must not be silently dropped');
        assert.strictEqual(comments[0].text, 'Verify this cell value.');
        assert.strictEqual(comments[0].author, 'Alice');
        assert.strictEqual(comments[0].anchorText, 'Cell text');
    });

    await test('accumulates anchor text across paragraph boundary', async () => {
        const buf = await buildDocx(MULTI_PARA_DOCUMENT_XML, MULTI_PARA_COMMENTS_XML);
        const ast = await OfficeParser.parseOffice(buf);
        const comments = (ast.blocks ?? []).filter(b => b.type === 'comment') as CommentBlock[];

        assert.strictEqual(comments.length, 1, 'multi-paragraph comment must be emitted');
        assert.strictEqual(comments[0].text, 'Spans two paragraphs.');
        assert.strictEqual(comments[0].author, 'Bob');
        assert.ok(
            comments[0].anchorText?.includes('First paragraph text'),
            `anchorText should include first paragraph; got: ${comments[0].anchorText}`
        );
        assert.ok(
            comments[0].anchorText?.includes('Second paragraph text'),
            `anchorText should include second paragraph; got: ${comments[0].anchorText}`
        );
    });

    console.log(`\n${passed} passed, ${failed} failed\n`);
    if (failed > 0) process.exit(1);
}

runTests().catch(e => { console.error(e); process.exit(1); });
