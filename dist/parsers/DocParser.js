"use strict";
/**
 * Legacy Word Document (.doc) Parser
 *
 * Handles the legacy binary Compound Document File Format (CFB) used by Word 97-2003.
 * Uses the `word-extractor` package to extract body text and comment annotations.
 *
 * Limitations vs. DOCX:
 * - Comment author, date, anchor text, and reply threading are not available from
 *   the binary format via word-extractor. Comment text is returned as a flat string.
 * - Tracked changes (insertions/deletions) are not exposed by word-extractor.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseDoc = void 0;
const parseDoc = async (buffer, config) => {
    // word-extractor is a CommonJS module
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const WordExtractor = require('word-extractor');
    const extractor = new WordExtractor();
    const doc = await extractor.extract(buffer);
    const bodyText = doc.getBody() ?? '';
    const annotationsText = config.ignoreComments ? '' : (doc.getAnnotations() ?? '');
    const blocks = [];
    if (bodyText.trim()) {
        const textBlock = { type: 'text', content: bodyText };
        blocks.push(textBlock);
    }
    if (annotationsText.trim()) {
        const commentBlock = {
            type: 'comment',
            text: annotationsText.trim(),
        };
        blocks.push(commentBlock);
    }
    const fullText = bodyText;
    return {
        type: 'doc',
        metadata: {},
        content: [],
        attachments: [],
        fullText,
        blocks,
        images: [],
        toText: () => fullText,
    };
};
exports.parseDoc = parseDoc;
