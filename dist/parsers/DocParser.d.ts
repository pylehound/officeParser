/**
 * Legacy Word Document (.doc) Parser
 *
 * Handles the legacy binary Compound Document File Format (CFB) used by Word 97-2003.
 * Uses the `word-extractor` package to extract body text.
 * Uses `cfb` to parse ATRD + GrpXstAtnOwners + PlcfandTxt for per-annotation author info.
 *
 * Limitations vs. DOCX:
 * - Comment anchor text and reply threading are not available.
 * - Dates are not available for Word 97 (nFib=0xC1); Word 2002+ required for ATRDPOST10.
 * - Tracked changes (insertions/deletions) are not exposed by word-extractor.
 */
/// <reference types="node" />
import { OfficeParserAST, OfficeParserConfig } from '../types';
export declare const parseDoc: (buffer: Buffer, config: OfficeParserConfig) => Promise<OfficeParserAST>;
