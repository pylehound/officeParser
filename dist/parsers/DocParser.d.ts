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
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
export declare const parseDoc: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;
