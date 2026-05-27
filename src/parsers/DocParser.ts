/**
 * Legacy Word Document (.doc) Parser
 *
 * Handles the legacy binary Compound Document File Format (CFB) used by Word 97-2003.
 * Uses the `word-extractor` package to extract body text.
 * Uses `cfb` to parse ATRD + GrpXstAtnOwners + PlcfandTxt for per-annotation author info.
 *
 * Limitations vs. DOCX:
 * - Comment anchor text, reply threading, and per-comment dates are not available.
 * - Tracked changes (insertions/deletions) are not exposed by word-extractor.
 */

import CFB from 'cfb';
import { Block, CommentBlock, OfficeParserAST, OfficeParserConfig, TextBlock } from '../types';
import { logWarning } from '../utils/errorUtils';

interface AnnotationInfo {
    author: string | undefined;
    text: string;
}

function parseXstOwners(buf: Buffer, offset: number, lcb: number): string[] {
    const authors: string[] = [];
    let off = offset;
    const end = offset + lcb;
    while (off + 2 <= end) {
        const cch = buf.readUInt16LE(off);
        off += 2;
        if (cch === 0 || off + cch * 2 > end) break;
        authors.push(buf.slice(off, off + cch * 2).toString('ucs2'));
        off += cch * 2;
        // XstZ format: each string is followed by a 2-byte null terminator
        if (off + 2 <= end && buf.readUInt16LE(off) === 0) off += 2;
    }
    return authors;
}

function parseAnnotationMetadata(buffer: Buffer, annotationsText: string, config: OfficeParserConfig): AnnotationInfo[] | null {
    try {
        const cfb = CFB.read(buffer, { type: 'buffer' });
        const wdEntry = CFB.find(cfb, 'WordDocument');
        if (!wdEntry) return null;
        const wdBuf = Buffer.from(wdEntry.content as Buffer);

        // Determine which table stream (fWhichTblStm bit 9 of FibBase.flags at 0x0A)
        const fibFlags = wdBuf.readUInt16LE(0x0A);
        const tblName = (fibFlags & 0x0200) !== 0 ? '1Table' : '0Table';
        const tblEntry = CFB.find(cfb, tblName);
        if (!tblEntry) return null;
        const tblBuf = Buffer.from(tblEntry.content as Buffer);

        const fcPlcfandRef  = wdBuf.readInt32LE(0xBA);
        const lcbPlcfandRef = wdBuf.readUInt32LE(0xBE);
        const fcPlcfandTxt  = wdBuf.readInt32LE(0xC2);
        const lcbPlcfandTxt = wdBuf.readUInt32LE(0xC6);
        const fcGrpOwners   = wdBuf.readInt32LE(0x16A);
        const lcbGrpOwners  = wdBuf.readUInt32LE(0x16E);
        const ccpText       = wdBuf.readUInt32LE(0x4C);

        if (lcbPlcfandRef < 28 || lcbPlcfandTxt < 8) return null;

        // PlcfandRef: (nAtn+1)*4 bytes of CPs + nAtn*20 bytes of ATRD
        const nAtn = (lcbPlcfandRef - 4) / 24;
        if (!Number.isInteger(nAtn) || nAtn < 1) return null;
        if (lcbPlcfandTxt !== (nAtn + 1) * 4) return null;

        // Validate buffer bounds
        if (fcPlcfandRef < 0 || fcPlcfandRef + lcbPlcfandRef > tblBuf.length) return null;
        if (fcPlcfandTxt < 0 || fcPlcfandTxt + lcbPlcfandTxt > tblBuf.length) return null;

        const authors = lcbGrpOwners > 0 && fcGrpOwners >= 0 && fcGrpOwners + lcbGrpOwners <= tblBuf.length
            ? parseXstOwners(tblBuf, fcGrpOwners, lcbGrpOwners)
            : [];

        const atrdStart = fcPlcfandRef + (nAtn + 1) * 4;
        if (atrdStart + nAtn * 20 > tblBuf.length) return null;
        const result: AnnotationInfo[] = [];

        for (let i = 0; i < nAtn; i++) {
            const cpStart  = tblBuf.readUInt32LE(fcPlcfandTxt + i * 4);
            const cpEnd    = tblBuf.readUInt32LE(fcPlcfandTxt + (i + 1) * 4);
            const relStart = cpStart - ccpText;
            // word-extractor may strip the trailing \r of the last annotation,
            // so clamp rather than reject when relEnd slightly overshoots.
            const relEnd   = Math.min(cpEnd - ccpText, annotationsText.length);
            if (relStart < 0 || relStart >= relEnd) return null;

            const text = annotationsText.substring(relStart, relEnd).replace(/\r+$/, '').trim();
            const ibst  = tblBuf.readUInt16LE(atrdStart + i * 20 + 10);
            const author = authors[ibst];
            result.push({ author, text });
        }

        return result;
    } catch (error) {
        logWarning('Failed to parse .doc annotation metadata:', config, error);
        return null;
    }
}

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
        const perAnnotation = parseAnnotationMetadata(buffer, annotationsText, config);
        if (perAnnotation) {
            for (const { author, text } of perAnnotation) {
                if (text) {
                    const commentBlock: CommentBlock = { type: 'comment', text, author };
                    blocks.push(commentBlock);
                }
            }
        } else {
            const commentBlock: CommentBlock = { type: 'comment', text: annotationsText.trim() };
            blocks.push(commentBlock);
        }
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
