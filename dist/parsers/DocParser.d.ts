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
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
export declare const parseDoc: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;
