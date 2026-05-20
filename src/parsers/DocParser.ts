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

import { Block, CommentBlock, OfficeParserAST, OfficeParserConfig, TextBlock } from '../types';

export const parseDoc = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    // word-extractor is a CommonJS module
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const WordExtractor = require('word-extractor');
    const extractor = new WordExtractor();
    const doc = await extractor.extract(buffer);

    const bodyText: string = doc.getBody() ?? '';
    const annotationsText: string = config.ignoreComments ? '' : (doc.getAnnotations() ?? '');

    const blocks: Block[] = [];

    if (bodyText.trim()) {
        const textBlock: TextBlock = { type: 'text', content: bodyText };
        blocks.push(textBlock);
    }

    if (annotationsText.trim()) {
        const commentBlock: CommentBlock = {
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
