/**
 * Word Document (DOCX) Parser
 * 
 * **DOCX Format Overview:**
 * DOCX is the default format for Microsoft Word documents since Office 2007.
 * It's based on the Office Open XML (OOXML) standard (ECMA-376, ISO/IEC 29500).
 * 
 * **File Structure:**
 * DOCX files are ZIP archives containing:
 * - `word/document.xml` - Main document content
 * - `word/styles.xml` - Style definitions
 * - `word/numbering.xml` - List numbering definitions
 * - `word/footnotes.xml` - Footnotes content
 * - `word/media/*` - Embedded images and media
 * - `docProps/core.xml` - Document metadata
 * - `[Content_Types].xml` - MIME type mappings
 * 
 * **XML Structure (word/document.xml):**
 * ```xml
 * <w:document>
 *   <w:body>
 *     <w:p>                    <!-- Paragraph -->
 *       <w:pPr>                <!-- Paragraph properties -->
 *         <w:pStyle w:val="Heading1"/>
 *       </w:pPr>
 *       <w:r>                  <!-- Run (text with same formatting) -->
 *         <w:rPr>              <!-- Run properties -->
 *           <w:b/>             <!-- Bold -->
 *           <w:sz w:val="24"/> <!-- Font size (half-points) -->
 *         </w:rPr>
 *         <w:t>Hello</w:t>     <!-- Text -->
 *       </w:r>
 *     </w:p>
 *   </w:body>
 * </w:document>
 * ```
 * 
 * **Key OOXML Elements:**
 * - `<w:p>` - Paragraph
 * - `<w:r>` - Run (contiguous text with same formatting)
 * - `<w:t>` - Text content
 * - `<w:b>`, `<w:i>`, `<w:u>` - Bold, italic, underline
 * - `<w:pStyle>` - Paragraph style (for headings)
 * - `<w:numPr>` - List numbering properties
 * - `<w:tbl>` - Table
 * - `<w:drawing>` - Drawing/image
 * 
 * **Parsing Approach:**
 * 1. Extract ZIP contents
 * 2. Parse word/document.xml for structure and text
 * 3. Extract formatting from run properties (rPr)
 * 4. Identify headings via paragraph styles
 * 5. Extract footnotes from word/footnotes.xml
 * 6. Process embedded images from word/media/*
 * 7. Parse metadata from docProps/core.xml
 * 
 * @module WordParser
 * @see https://www.ecma-international.org/publications-and-standards/standards/ecma-376/ OOXML Standard
 * @see https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/ [MS-DOCX] Specification
 */

import { XMLSerializer } from '@xmldom/xmldom';
import { Block, ChartBlock, ChartMetadata, CommentBlock, CommentMetadata, CommentReply, DeletionBlock, FooterBlock, HeaderBlock, ImageBlock, ImageMetadata, InsertionBlock, ListMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeParserConfig, TableBlock, TextBlock, TextFormatting, TextMetadata, TrackedChangeMetadata } from '../types';
import { logWarning } from '../utils/errorUtils';
import { createAttachment } from '../utils/imageUtils';
import { performOcr } from '../utils/ocrUtils';
import { getDirectChildren, getElementsByTagName, parseOfficeMetadata, parseXmlString } from '../utils/xmlUtils';
import { extractFiles } from '../utils/zipUtils';

/**
 * Parses a Word document (.docx) and extracts content, formatting, and metadata.
 * 
 * The parsing process:
 * 1. Unzip the DOCX file
 * 2. Parse word/document.xml to extract paragraphs and runs
 * 3. Extract text formatting from run properties
 * 4. Identify headings from paragraph styles
 * 5. Process lists from numbering properties
 * 6. Extract images and optionally perform OCR
 * 7. Parse document metadata
 * 
 * @param buffer - The DOCX file as a Buffer
 * @param config - Parser configuration options
 * @returns A promise resolving to the parsed AST
 */
export const parseWord = async (buffer: Buffer, config: OfficeParserConfig): Promise<OfficeParserAST> => {
    const documentFileRegex = /word\/document[\d+]?.xml/;
    const footnotesFileRegex = /word\/footnotes[\d+]?.xml/;
    const endnotesFileRegex = /word\/endnotes[\d+]?.xml/;
    const numberingFileRegex = /word\/numbering[\d+]?.xml/;
    const mediaFileRegex = /(word\/)?media\/.*/;
    const corePropsFileRegex = /docProps\/core[\d+]?.xml/;
    const relsFileRegex = /word\/_rels\/document[\d+]?.xml\.rels/;
    const stylesFileRegex = /word\/styles[\d+]?.xml/;
    const commentsFileRegex = /word\/comments[\d+]?.xml/;
    const commentsExtendedFileRegex = /word\/commentsExtended[\d+]?.xml/;
    const headerFileRegex = /word\/header[\d+]?.xml/;
    const footerFileRegex = /word\/footer[\d+]?.xml/;

    const xmlSerializer = new XMLSerializer();

    // Pre-compiled regexes for run-property boolean tags (used in hot path)
    const REGEX_W_B = /<w:b(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const REGEX_W_I = /<w:i(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const REGEX_W_STRIKE = /<w:strike(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const REGEX_W_DSTRIKE = /<w:dstrike(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const getBoolValFromRegex = (xmlSnippet: string, regex: RegExp): boolean | null => {
        const match = xmlSnippet.match(regex);
        if (match) {
            const val = match[1];
            if (val === undefined) return true;
            return val === '1' || val === 'true' || val === 'on';
        }
        return null;
    };

    // Helper to extract formatting from run properties XML string
    const extractFormattingFromXml = (rPr: Element): TextFormatting => {
        const formatting: TextFormatting = {};
        const rPrString = xmlSerializer.serializeToString(rPr);

        const bold = getBoolValFromRegex(rPrString, REGEX_W_B);
        if (bold !== null) formatting.bold = bold;

        const italic = getBoolValFromRegex(rPrString, REGEX_W_I);
        if (italic !== null) formatting.italic = italic;

        const underlineMatch = rPrString.match(/<w:u(?: w:val="([^"]+)")?\/?>/);
        if (underlineMatch) {
            const val = underlineMatch[1];
            // If val is missing, it's a default underline (true). 
            // If val is present, it's true unless explicit 'none'.
            if (!val || val !== 'none') {
                formatting.underline = true;
            }
        }

        const strike = getBoolValFromRegex(rPrString, REGEX_W_STRIKE);
        const dstrike = getBoolValFromRegex(rPrString, REGEX_W_DSTRIKE);
        if (strike !== null) formatting.strikethrough = strike;
        else if (dstrike !== null) formatting.strikethrough = dstrike;

        // Font size
        const szMatch = rPrString.match(/<w:sz w:val="(\d+)"/);
        if (szMatch) formatting.size = (parseInt(szMatch[1]) / 2).toString() + 'pt';

        // Color
        const colorMatch = rPrString.match(/<w:color w:val="([^"]+)"/);
        if (colorMatch && colorMatch[1] !== 'auto') formatting.color = '#' + colorMatch[1];

        // Background color (shading)
        const shdMatch = rPrString.match(/<w:shd[^>]*w:fill="([^"]+)"/);
        if (shdMatch && shdMatch[1] !== 'auto') formatting.backgroundColor = '#' + shdMatch[1];

        // Highlight (map to backgroundColor)
        const highlightMatch = rPrString.match(/<w:highlight w:val="([^"]+)"/);
        if (highlightMatch && highlightMatch[1] !== 'none') {
            const colorMap: { [key: string]: string } = {
                'yellow': '#FFFF00', 'green': '#00FF00', 'cyan': '#00FFFF', 'magenta': '#FF00FF',
                'blue': '#0000FF', 'red': '#FF0000', 'darkBlue': '#00008B', 'darkCyan': '#008B8B',
                'darkGreen': '#006400', 'darkMagenta': '#8B008B', 'darkRed': '#8B0000',
                'darkYellow': '#808000', 'darkGray': '#A9A9A9', 'lightGray': '#D3D3D3', 'black': '#000000'
            };
            formatting.backgroundColor = colorMap[highlightMatch[1]] || highlightMatch[1];
        }

        // Font family
        const rFontsMatch = rPrString.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/);
        if (rFontsMatch) {
            formatting.font = rFontsMatch[1];
        } else {
            const hAnsiMatch = rPrString.match(/<w:rFonts[^>]*w:hAnsi="([^"]+)"/);
            if (hAnsiMatch) formatting.font = hAnsiMatch[1];
        }

        // Subscript/Superscript
        const vertAlignMatch = rPrString.match(/<w:vertAlign w:val="([^"]+)"/);
        if (vertAlignMatch) {
            if (vertAlignMatch[1] === 'subscript') formatting.subscript = true;
            if (vertAlignMatch[1] === 'superscript') formatting.superscript = true;
        }

        return formatting;
    };

    const files = await extractFiles(buffer, x =>
        !!x.match(documentFileRegex) ||
        !!x.match(footnotesFileRegex) ||
        !!x.match(endnotesFileRegex) ||
        !!x.match(numberingFileRegex) ||
        !!x.match(corePropsFileRegex) ||
        !!x.match(relsFileRegex) ||
        !!x.match(stylesFileRegex) ||
        !!x.match(commentsFileRegex) ||
        !!x.match(commentsExtendedFileRegex) ||
        !!x.match(headerFileRegex) ||
        !!x.match(footerFileRegex) ||
        (!!config.extractAttachments && !!x.match(mediaFileRegex))
    );

    type FileEntry = (typeof files)[0];
    let corePropsFile: FileEntry | undefined;
    let relsFile: FileEntry | undefined;
    let numberingFile: FileEntry | undefined;
    let stylesFile: FileEntry | undefined;
    let footnotesFile: FileEntry | undefined;
    let endnotesFile: FileEntry | undefined;
    let commentsFile: FileEntry | undefined;
    let commentsExtendedFile: FileEntry | undefined;
    const headerFiles: FileEntry[] = [];
    const footerFiles: FileEntry[] = [];
    for (const f of files) {
        if (f.path.match(corePropsFileRegex)) corePropsFile = f;
        else if (f.path.match(relsFileRegex)) relsFile = f;
        else if (f.path.match(numberingFileRegex)) numberingFile = f;
        else if (f.path.match(stylesFileRegex)) stylesFile = f;
        else if (f.path.match(footnotesFileRegex)) footnotesFile = f;
        else if (f.path.match(endnotesFileRegex)) endnotesFile = f;
        else if (f.path.match(commentsExtendedFileRegex)) commentsExtendedFile = f;
        else if (f.path.match(commentsFileRegex)) commentsFile = f;
        else if (f.path.match(headerFileRegex)) headerFiles.push(f);
        else if (f.path.match(footerFileRegex)) footerFiles.push(f);
    }

    // Extract metadata
    const metadata = corePropsFile ? parseOfficeMetadata(corePropsFile.content.toString()) : {};

    const footnoteMap = new Map<string, OfficeContentNode[]>();
    const endnoteMap = new Map<string, OfficeContentNode[]>();
    const collectedNotes: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];
    const attachmentByNameAndType = new Map<string, OfficeAttachment>();
    const mediaFiles = files.filter(f => f.path.match(mediaFileRegex));

    // Extract relationships
    const relsMap: { [key: string]: string } = {};
    if (relsFile) {
        const relsXml = parseXmlString(relsFile.content.toString());
        const relationships = getElementsByTagName(relsXml, "Relationship");
        for (let i = 0; i < relationships.length; i++) {
            const id = relationships[i].getAttribute("Id");
            const target = relationships[i].getAttribute("Target");
            if (id && target) {
                relsMap[id] = target;
            }
        }
    }

    const numberingMap: { [key: string]: { [key: string]: { numFmt: string, lvlText: string } } } = {};

    if (numberingFile) {
        const numberingXml = parseXmlString(numberingFile.content.toString());
        const nums = getElementsByTagName(numberingXml, "w:num");
        const abstractNums = getElementsByTagName(numberingXml, "w:abstractNum");

        const abstractNumMap: { [key: string]: any } = {};
        for (let i = 0; i < abstractNums.length; i++) {
            const abstractNumId = abstractNums[i].getAttribute("w:abstractNumId");
            if (abstractNumId) {
                abstractNumMap[abstractNumId] = abstractNums[i];
            }
        }

        for (let i = 0; i < nums.length; i++) {
            const numId = nums[i].getAttribute("w:numId");
            const abstractNumIdNode = getElementsByTagName(nums[i], "w:abstractNumId")[0];
            const abstractNumId = abstractNumIdNode?.getAttribute("w:val");

            if (numId && abstractNumId && abstractNumMap[abstractNumId]) {
                numberingMap[numId] = {};
                const lvls = getElementsByTagName(abstractNumMap[abstractNumId], "w:lvl");
                for (let j = 0; j < lvls.length; j++) {
                    const ilvl = lvls[j].getAttribute("w:ilvl");
                    const numFmtNode = getElementsByTagName(lvls[j], "w:numFmt")[0];
                    const lvlTextNode = getElementsByTagName(lvls[j], "w:lvlText")[0];
                    if (ilvl) {
                        numberingMap[numId][ilvl] = {
                            numFmt: numFmtNode?.getAttribute("w:val") || 'decimal',
                            lvlText: lvlTextNode?.getAttribute("w:val") || ''
                        };
                    }
                }
            }
        }
    }

    // Parse Styles once and derive styleMap, docDefaults, defaultParaStyleId
    const styleMap: { [key: string]: { formatting: TextFormatting, alignment?: 'left' | 'center' | 'right' | 'justify', backgroundColor?: string } } = {};
    let docDefaults: Partial<TextFormatting> = {};
    let defaultParaStyleId: string | undefined = undefined;

    if (stylesFile) {
        const stylesXml = parseXmlString(stylesFile.content.toString());
        const styles = getElementsByTagName(stylesXml, "w:style");

        for (let i = 0; i < styles.length; i++) {
            const styleId = styles[i].getAttribute("w:styleId");
            if (styleId) {
                const rPr = getElementsByTagName(styles[i], "w:rPr")[0];
                const pPr = getElementsByTagName(styles[i], "w:pPr")[0];

                const formatting = rPr ? extractFormattingFromXml(rPr) : {};
                let alignment: 'left' | 'center' | 'right' | 'justify' | undefined = undefined;
                let backgroundColor: string | undefined = undefined;

                if (pPr) {
                    const jc = getElementsByTagName(pPr, "w:jc")[0];
                    if (jc) {
                        const val = jc.getAttribute("w:val");
                        if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                            alignment = val;
                        }
                    }
                    const shd = getElementsByTagName(pPr, "w:shd")[0];
                    if (shd) {
                        const fill = shd.getAttribute("w:fill");
                        if (fill && fill !== 'auto') backgroundColor = '#' + fill;
                    }
                }

                styleMap[styleId] = { formatting, alignment, backgroundColor };

                // Detect default paragraph style (w:type="paragraph" and w:default="1")
                const styleType = styles[i].getAttribute("w:type");
                const isDefault = styles[i].getAttribute("w:default");
                if (styleType === "paragraph" && isDefault === "1" && !defaultParaStyleId) {
                    defaultParaStyleId = styleId;
                }
            }
        }

        const docDefaultsNode = getElementsByTagName(stylesXml, "w:docDefaults")[0];
        if (docDefaultsNode) {
            const rPrDefaultNode = getElementsByTagName(docDefaultsNode, "w:rPrDefault")[0];
            if (rPrDefaultNode) {
                const rPr = getElementsByTagName(rPrDefaultNode, "w:rPr")[0];
                if (rPr) {
                    docDefaults = extractFormattingFromXml(rPr);
                }
            }
        }

        if (!defaultParaStyleId && styleMap["Normal"]) {
            defaultParaStyleId = "Normal";
        }
    }



    // Parse comments from word/comments.xml
    interface StructuredComment {
        id: string;
        author?: string;
        date?: string;
        text: string;
        replies: CommentReply[];
    }
    const structuredCommentMap = new Map<string, StructuredComment>();

    // Build paraId -> parentParaId map from commentsExtended.xml as a fallback
    // for documents where w14:paraIdParent is absent on comment paragraphs
    const extendedParentMap = new Map<string, string>();
    if (commentsExtendedFile && !config.ignoreComments) {
        const extDoc = parseXmlString(commentsExtendedFile.content.toString());
        // getElementsByTagName matches on the serialised prefix literal, not namespace URI.
        // Documents that bind the same namespace to a different prefix would be missed here.
        const commentExNodes = getElementsByTagName(extDoc, "w15:commentEx");
        for (const node of commentExNodes) {
            const paraId = node.getAttribute("w15:paraId");
            const parentParaId = node.getAttribute("w15:paraIdParent");
            if (paraId && parentParaId) {
                extendedParentMap.set(paraId, parentParaId);
            }
        }
    }

    if (commentsFile && !config.ignoreComments) {
        const commentsDoc = parseXmlString(commentsFile.content.toString());
        const commentNodes = getElementsByTagName(commentsDoc, "w:comment");

        interface RawComment {
            id: string;
            author?: string;
            date?: string;
            text: string;
            firstParaId?: string;
            parentParaId?: string;
        }
        const rawComments: RawComment[] = [];
        const paraIdToCommentId = new Map<string, string>();

        for (const commentNode of commentNodes) {
            const id = commentNode.getAttribute("w:id");
            if (!id) continue;

            const author = commentNode.getAttribute("w:author") ?? undefined;
            const date = commentNode.getAttribute("w:date") ?? undefined;

            const pNodes = getElementsByTagName(commentNode, "w:p");
            const firstPara = pNodes[0];

            // w14:paraId and w14:paraIdParent use the w14 namespace
            const firstParaId = firstPara?.getAttribute("w14:paraId") ?? undefined;
            const parentParaId = firstPara?.getAttribute("w14:paraIdParent")
                || (firstParaId ? extendedParentMap.get(firstParaId) : undefined);

            const text = pNodes
                .map(p => getElementsByTagName(p, "w:t").map(t => t.textContent ?? '').join(''))
                .filter(t => t)
                .join('\n');

            rawComments.push({ id, author, date, text, firstParaId, parentParaId });
            if (firstParaId) paraIdToCommentId.set(firstParaId, id);
        }

        // Build top-level comments first
        for (const raw of rawComments) {
            if (!raw.parentParaId) {
                structuredCommentMap.set(raw.id, {
                    id: raw.id, author: raw.author, date: raw.date, text: raw.text, replies: []
                });
            }
        }

        // Attach replies to their parent comment
        for (const raw of rawComments) {
            if (raw.parentParaId) {
                const parentId = paraIdToCommentId.get(raw.parentParaId);
                if (parentId) {
                    const parent = structuredCommentMap.get(parentId);
                    if (parent) {
                        parent.replies.push({ text: raw.text, author: raw.author, date: raw.date });
                    }
                } else {
                    // Parent not found — treat as top-level
                    structuredCommentMap.set(raw.id, {
                        id: raw.id, author: raw.author, date: raw.date, text: raw.text, replies: []
                    });
                }
            }
        }
    }

    // Tracks which comment IDs were referenced in the current paragraph (populated during parseParagraph)
    let pendingCommentRefs: string[] = [];
    // Tracked-change nodes collected during paragraph parsing, emitted as top-level content after the paragraph
    let pendingTrackedChanges: OfficeContentNode[] = [];
    // Anchor text accumulated between commentRangeStart and commentRangeEnd
    const commentAnchorAccum = new Map<string, string>();
    // Open comment ranges that span paragraph boundaries (keyed by comment ID)
    const activeCommentRanges = new Set<string>();

    const content: OfficeContentNode[] = [];
    const rawContents: string[] = [];
    const numberingState: { [key: string]: { [key: string]: number } } = {};
    const listCounters: { [key: string]: { [key: string]: number } } = {}; // Track item index per listId/level

    // Helper to parse a paragraph node
    const parseParagraph = (pNode: Element, isNoteContext = false, skipReset = false): OfficeContentNode => {
        const pXml = xmlSerializer.serializeToString(pNode);

        // Check if it's a list item
        const numPr = getElementsByTagName(pNode, "w:numPr")[0];
        const isList = !!numPr;

        // Check if it's a heading
        const pPr = getElementsByTagName(pNode, "w:pPr")[0];
        const pStyle = pPr ? getElementsByTagName(pPr, "w:pStyle")[0] : null;
        const pStyleVal = pStyle ? pStyle.getAttribute("w:val") : null;
        const isHeading = pStyleVal ? (pStyleVal.startsWith("Heading") || pStyleVal === "Title") : false;

        // Extract Paragraph Style Properties
        const styleProps = pStyleVal && styleMap[pStyleVal] ? styleMap[pStyleVal] : { formatting: {} };

        // Extract Alignment
        let alignment = styleProps.alignment;
        if (pPr) {
            const jc = getElementsByTagName(pPr, "w:jc")[0];
            if (jc) {
                const val = jc.getAttribute("w:val");
                if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                    alignment = val;
                }
            }
        }

        // Extract Paragraph Background
        let paraBackgroundColor = styleProps.backgroundColor;
        if (pPr) {
            const shd = getElementsByTagName(pPr, "w:shd")[0];
            if (shd) {
                const fill = shd.getAttribute("w:fill");
                if (fill && fill !== 'auto') {
                    paraBackgroundColor = '#' + fill;
                }
            }
        }

        // Extract paragraph-level run properties
        let paragraphRunFormatting: TextFormatting = { ...styleProps.formatting };
        if (pPr) {
            const pPrRPr = getElementsByTagName(pPr, "w:rPr")[0];
            if (pPrRPr) {
                const pPrFormatting = extractFormattingFromXml(pPrRPr);
                for (const key in pPrFormatting) {
                    const value = pPrFormatting[key as keyof TextFormatting];
                    if (value === false) {
                        delete paragraphRunFormatting[key as keyof TextFormatting];
                    } else if (value !== undefined) {
                        (paragraphRunFormatting as any)[key] = value;
                    }
                }
            }
        }

        // Extract text and children
        let text = '';
        const children: OfficeContentNode[] = [];

        // Reset pending state for a fresh top-level paragraph. Skip the reset when recursing
        // into text-box content (skipReset=true) so refs collected before the text-box run are
        // preserved, and skip for note contexts (isNoteContext=true) which have their own domain.
        if (!isNoteContext && !skipReset) {
            pendingCommentRefs = [];
            pendingTrackedChanges = [];
        }

        // Traverse children of paragraph (runs, hyperlinks, etc.)
        const processChildNode = (node: Node) => {
            if (node.nodeName === 'w:commentRangeStart' && !config.ignoreComments) {
                const id = (node as Element).getAttribute('w:id');
                if (id) {
                    activeCommentRanges.add(id);
                    if (!commentAnchorAccum.has(id)) commentAnchorAccum.set(id, '');
                }
            } else if (node.nodeName === 'w:commentRangeEnd' && !config.ignoreComments) {
                const id = (node as Element).getAttribute('w:id');
                if (id) activeCommentRanges.delete(id);
            } else if (node.nodeName === 'w:ins' && !config.ignoreTrackedChanges) {
                const insNode = node as Element;
                const author = insNode.getAttribute('w:author') ?? undefined;
                const date = insNode.getAttribute('w:date') ?? undefined;
                const runs = getElementsByTagName(insNode, 'w:t');
                const insertedText = runs.map(r => r.textContent ?? '').join('');
                if (insertedText) {
                    for (const id of activeCommentRanges) {
                        commentAnchorAccum.set(id, (commentAnchorAccum.get(id) ?? '') + insertedText);
                    }
                    // Emit as a top-level node after the paragraph (side-channel, not a child)
                    if (!isNoteContext) {
                        pendingTrackedChanges.push({
                            type: 'insertion',
                            text: insertedText,
                            metadata: { author, date } as TrackedChangeMetadata
                        });
                    }
                }
            } else if (node.nodeName === 'w:del' && !config.ignoreTrackedChanges) {
                const delNode = node as Element;
                const author = delNode.getAttribute('w:author') ?? undefined;
                const date = delNode.getAttribute('w:date') ?? undefined;
                // Deleted text uses w:delText, not w:t
                const runs = getElementsByTagName(delNode, 'w:delText');
                const deletedText = runs.map(r => r.textContent ?? '').join('');
                if (deletedText && !isNoteContext) {
                    pendingTrackedChanges.push({
                        type: 'deletion',
                        text: deletedText,
                        metadata: { author, date } as TrackedChangeMetadata
                    });
                }
            } else if (node.nodeName === 'w:r') {
                const runNode = node as Element;
                const rPr = getElementsByTagName(runNode, "w:rPr")[0];

                // Formatting
                let formatting: TextFormatting = {};
                // Apply paragraph-level formatting
                for (const key in paragraphRunFormatting) {
                    (formatting as any)[key] = (paragraphRunFormatting as any)[key];
                }

                // Check for run style
                const rStyle = rPr ? getElementsByTagName(rPr, "w:rStyle")[0] : null;
                const rStyleVal = rStyle ? rStyle.getAttribute("w:val") : pStyleVal;
                if (rStyleVal && styleMap[rStyleVal]) {
                    for (const key in styleMap[rStyleVal].formatting) {
                        (formatting as any)[key] = (styleMap[rStyleVal].formatting as any)[key];
                    }
                }

                // Apply direct run properties
                if (rPr) {
                    const directFormatting = extractFormattingFromXml(rPr);
                    for (const key in directFormatting) {
                        const value = directFormatting[key as keyof TextFormatting];
                        if (value === false) {
                            delete formatting[key as keyof TextFormatting];
                        } else if (value !== undefined) {
                            formatting[key as keyof TextFormatting] = value as any;
                        }
                    }
                }

                // Inherit paragraph background
                if (!formatting.backgroundColor && paraBackgroundColor) {
                    formatting.backgroundColor = paraBackgroundColor;
                }

                // Text content
                const tNodes = getElementsByTagName(runNode, "w:t");
                for (const tNode of tNodes) {
                    const tContent = tNode.textContent || '';
                    text += tContent;
                    // Accumulate anchor text for any open comment ranges
                    for (const id of activeCommentRanges) {
                        commentAnchorAccum.set(id, (commentAnchorAccum.get(id) ?? '') + tContent);
                    }
                    const textNode: OfficeContentNode = {
                        type: 'text',
                        text: tContent,
                        formatting: formatting
                    };
                    if (config.includeRawContent) {
                        textNode.rawContent = xmlSerializer.serializeToString(tNode);
                    }
                    // Always set a style: run style > paragraph style > detected default
                    // Use detected default style for international compatibility
                    const nodeStyle = rStyleVal || pStyleVal || defaultParaStyleId;
                    if (nodeStyle) {
                        textNode.metadata = { style: nodeStyle };
                    }
                    children.push(textNode);
                }

                // Comment references
                if (!config.ignoreComments && !isNoteContext) {
                    const commentRef = getElementsByTagName(runNode, "w:commentReference")[0];
                    if (commentRef) {
                        const id = commentRef.getAttribute("w:id");
                        if (id && structuredCommentMap.has(id)) {
                            pendingCommentRefs.push(id);
                        }
                    }
                }

                // Images/Drawings
                if (config.extractAttachments) {
                    const drawings = getElementsByTagName(runNode, "w:drawing");
                    const picts = getElementsByTagName(runNode, "w:pict");
                    const allImages = [...drawings, ...picts];

                    for (const imgNode of allImages) {
                        const imgXml = xmlSerializer.serializeToString(imgNode);

                        // Extract Alt Text
                        let altText = '';
                        const docPr = getElementsByTagName(imgNode, "wp:docPr")[0];
                        if (docPr) {
                            altText = docPr.getAttribute("descr") || docPr.getAttribute("title") || '';
                        }

                        // Extract Relationship ID
                        let rId = '';
                        const blip = getElementsByTagName(imgNode, "a:blip")[0];
                        if (blip) {
                            rId = blip.getAttribute("r:embed") || '';
                        } else {
                            const imagedata = getElementsByTagName(imgNode, "v:imagedata")[0];
                            if (imagedata) {
                                rId = imagedata.getAttribute("r:id") || '';
                            }
                        }

                        if (rId && relsMap[rId]) {
                            const target = relsMap[rId];
                            const filename = target.split('/').pop();
                            if (filename) {
                                const imageNode: OfficeContentNode = {
                                    type: 'image',
                                    text: '',
                                    metadata: { attachmentName: filename, altText: altText }
                                };
                                if (config.includeRawContent) {
                                    imageNode.rawContent = imgXml;
                                }
                                children.push(imageNode);
                            }
                        } else {
                            const imageNode: OfficeContentNode = {
                                type: 'image',
                                text: '',
                            };
                            if (config.includeRawContent) {
                                imageNode.rawContent = imgXml;
                            }
                            children.push(imageNode);
                        }
                    }
                }

                // Text boxes (modern: wps:txbx, legacy: v:textbox)
                // Text boxes: skipReset=true so refs/changes from before this run survive;
                // isNoteContext=false so comment refs and tracked changes are collected normally.
                const txbxContents = getElementsByTagName(runNode, "w:txbxContent");
                for (const txbx of txbxContents) {
                    const pNodes = getElementsByTagName(txbx, "w:p");
                    for (const p of pNodes) {
                        const parsed = parseParagraph(p as Element, false, true);
                        if (parsed.text?.trim()) children.push(parsed);
                    }
                }

                // Footnotes/Endnotes inside runs
                if (!config.ignoreNotes) {
                    const footnoteRef = getElementsByTagName(runNode, "w:footnoteReference")[0];
                    if (footnoteRef) {
                        const id = footnoteRef.getAttribute("w:id");
                        if (id && footnoteMap.has(id)) {
                            const noteNodes = footnoteMap.get(id)!;
                            const noteNode: OfficeContentNode = {
                                type: 'note',
                                text: noteNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'footnote', noteId: id }
                            } as any;

                            if (config.putNotesAtLast) {
                                collectedNotes.push(noteNode);
                            } else {
                                children.push(noteNode);
                            }
                        }
                    }

                    const endnoteRef = getElementsByTagName(runNode, "w:endnoteReference")[0];
                    if (endnoteRef) {
                        const id = endnoteRef.getAttribute("w:id");
                        if (id && endnoteMap.has(id)) {
                            const noteNodes = endnoteMap.get(id)!;
                            const noteNode: OfficeContentNode = {
                                type: 'note',
                                text: noteNodes.map((n: OfficeContentNode) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'endnote', noteId: id }
                            } as any;

                            if (config.putNotesAtLast) {
                                collectedNotes.push(noteNode);
                            } else {
                                children.push(noteNode);
                            }
                        }
                    }
                }
            } else if (node.nodeName === 'w:hyperlink') {
                const hlNode = node as Element;
                const rId = hlNode.getAttribute("r:id");
                const anchor = hlNode.getAttribute("w:anchor");

                let linkMetadata: TextMetadata | undefined;
                if (anchor) {
                    linkMetadata = { link: '#' + anchor, linkType: 'internal' };
                } else if (rId && relsMap[rId]) {
                    linkMetadata = { link: relsMap[rId], linkType: 'external' };
                }

                // Process children of hyperlink (usually runs)
                const hlChildren = Array.from(hlNode.childNodes);
                for (const child of hlChildren) {
                    // Capture the current length of children to apply metadata to new nodes
                    const startIndex = children.length;
                    processChildNode(child);
                    // Apply link metadata to the newly added text nodes
                    if (linkMetadata) {
                        for (let i = startIndex; i < children.length; i++) {
                            if (children[i].type === 'text') {
                                children[i].metadata = { ...(children[i].metadata ?? {}), ...linkMetadata };
                            }
                        }
                    }
                }
            }
        };

        const childNodes = Array.from(pNode.childNodes);
        for (const child of childNodes) {
            processChildNode(child);
        }

        if (isList) {
            const numIdNode = getElementsByTagName(numPr, "w:numId")[0];
            const ilvlNode = getElementsByTagName(numPr, "w:ilvl")[0];
            const numId = numIdNode ? numIdNode.getAttribute("w:val") || '0' : '0';
            const ilvl = ilvlNode ? parseInt(ilvlNode.getAttribute("w:val") || '0') : 0;

            let listType: 'ordered' | 'unordered' = 'ordered';
            let itemIndex = 0;
            if (numId && numberingMap[numId]) {
                const ilvlStr = ilvl.toString();
                if (!numberingState[numId]) numberingState[numId] = {};
                if (!numberingState[numId][ilvlStr]) numberingState[numId][ilvlStr] = 0;
                numberingState[numId][ilvlStr]++;
                for (let k = ilvl + 1; k < 10; k++) {
                    if (numberingState[numId][k.toString()]) numberingState[numId][k.toString()] = 0;
                }
                const numFmt = numberingMap[numId][ilvlStr]?.numFmt || 'decimal';
                listType = numFmt === 'bullet' ? 'unordered' : 'ordered';

                // Track itemIndex (starts at 0, continues across interruptions for same listId)
                if (!listCounters[numId]) listCounters[numId] = {};
                if (listCounters[numId][ilvlStr] === undefined) {
                    listCounters[numId][ilvlStr] = 0;
                } else {
                    listCounters[numId][ilvlStr]++;
                }
                itemIndex = listCounters[numId][ilvlStr];
            }

            const listNode: OfficeContentNode = {
                type: 'list',
                text: text,
                children: children,
                metadata: {
                    listType,
                    indentation: ilvl,
                    alignment: (alignment || 'left') as 'left' | 'center' | 'right' | 'justify',
                    listId: numId,
                    itemIndex: itemIndex,
                    style: pStyleVal
                } as ListMetadata
            };
            if (config.includeRawContent) listNode.rawContent = pXml;
            return listNode;

        } else if (isHeading) {
            const level = pStyleVal ? parseInt(pStyleVal.replace("Heading", "")) || 1 : 1;
            const headingNode: OfficeContentNode = {
                type: 'heading',
                text: text,
                children: children,
                metadata: { level, alignment, style: pStyleVal ?? undefined }
            };
            if (config.includeRawContent) headingNode.rawContent = pXml;
            return headingNode;
        } else {
            const paraNode: OfficeContentNode = {
                type: 'paragraph',
                text: text,
                children: children,
                metadata: { alignment, style: pStyleVal ?? undefined }
            };
            if (config.includeRawContent) paraNode.rawContent = pXml;
            return paraNode;
        }
    };

    // Drain pending comment refs and tracked-change nodes into a flat list and reset both arrays.
    // Called after every parseParagraph and parseTable invocation that may have populated them.
    const flushPendings = (): OfficeContentNode[] => {
        const flushed: OfficeContentNode[] = [];
        if (!config.ignoreTrackedChanges && pendingTrackedChanges.length > 0) {
            flushed.push(...pendingTrackedChanges);
            pendingTrackedChanges = [];
        }
        if (!config.ignoreComments && pendingCommentRefs.length > 0) {
            for (const commentId of pendingCommentRefs) {
                const comment = structuredCommentMap.get(commentId);
                if (comment) {
                    const anchorText = commentAnchorAccum.get(commentId);
                    commentAnchorAccum.delete(commentId);
                    const commentMeta: CommentMetadata = {
                        author: comment.author,
                        date: comment.date,
                        anchorText: anchorText ?? undefined,
                        replies: comment.replies.length > 0 ? comment.replies : undefined
                    };
                    flushed.push({ type: 'comment', text: comment.text, metadata: commentMeta });
                }
            }
            pendingCommentRefs = [];
        }
        return flushed;
    };

    // Helper to parse a table node
    const parseTable = (tblNode: Element): { node: OfficeContentNode; sidePending: OfficeContentNode[] } => {
        const rows: OfficeContentNode[] = [];
        const sidePending: OfficeContentNode[] = [];
        // Only get direct child rows, not nested table rows
        const trNodes = getDirectChildren(tblNode, "w:tr");

        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells: OfficeContentNode[] = [];
            // Only get direct child cells, not nested table cells
            const tcNodes = getDirectChildren(trNode, "w:tc");

            for (let cIndex = 0; cIndex < tcNodes.length; cIndex++) {
                const tcNode = tcNodes[cIndex];
                const cellChildren: OfficeContentNode[] = [];
                let cellText = '';

                // Cells contain paragraphs (and other block-level elements)
                const cellContentNodes = Array.from(tcNode.childNodes);
                for (const child of cellContentNodes) {
                    if (child.nodeName === 'w:p') {
                        const pNode = parseParagraph(child as Element);
                        cellChildren.push(pNode);
                        cellText += pNode.text;
                        // Flush any comments/tracked-changes from this cell paragraph
                        sidePending.push(...flushPendings());
                    } else if (child.nodeName === 'w:tbl') {
                        // Nested table — cascade its side-pending nodes up to ours
                        const nested = parseTable(child as Element);
                        cellChildren.push(nested.node);
                        sidePending.push(...nested.sidePending);
                        // Don't add nested table text to cell text - it will be handled recursively
                    }
                }

                const getCellText = (node: OfficeContentNode): string => {
                    let text = node.text || '';
                    if (node.children && node.children.length > 0) {
                        const childTexts = node.children.map(getCellText).filter(t => t !== '');
                        if (childTexts.length > 0) {
                            text += (text ? ' ' : '') + childTexts.join(' ');
                        }
                    }
                    return text;
                };

                const cellNode: OfficeContentNode = {
                    type: 'cell',
                    text: cellText,
                    children: cellChildren,
                    metadata: { row: rIndex, col: cIndex }
                };
                cellNode.text = getCellText(cellNode);
                cells.push(cellNode);
            }

            const rowNode: OfficeContentNode = {
                type: 'row',
                children: cells
            };
            rows.push(rowNode);
        }

        return {
            node: { type: 'table', children: rows },
            sidePending
        };
    };

    // Pre-process footnotes and endnotes to be inserted inline later
    if (!config.ignoreNotes) {
        if (footnotesFile) {
            const footnotesDoc = parseXmlString(footnotesFile.content.toString());
            const footnoteNodes = getElementsByTagName(footnotesDoc, "w:footnote");
            for (const node of footnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getElementsByTagName(node, "w:p");
                footnoteMap.set(id, pNodes.map(p => parseParagraph(p, true)));
            }
        }

        if (endnotesFile) {
            const endnotesDoc = parseXmlString(endnotesFile.content.toString());
            const endnoteNodes = getElementsByTagName(endnotesDoc, "w:endnote");
            for (const node of endnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0") continue;
                const pNodes = getElementsByTagName(node, "w:p");
                endnoteMap.set(id, pNodes.map(p => parseParagraph(p, true)));
            }
        }
    }

    // Extract headers and footers, deduplicated by plain text content, before body.
    const extractPlainText = (el: Element): string =>
        getElementsByTagName(el, "w:t").map(t => t.textContent || '').join('').trim();

    const seenHeaderTexts = new Set<string>();
    for (const hFile of headerFiles) {
        const hDoc = parseXmlString(hFile.content.toString());
        const hdrEl = getElementsByTagName(hDoc, "w:hdr")[0];
        if (!hdrEl) continue;
        const text = extractPlainText(hdrEl as Element);
        if (text && !seenHeaderTexts.has(text)) {
            seenHeaderTexts.add(text);
            content.push({ type: 'header', text } as OfficeContentNode);
        }
    }

    const seenFooterTexts = new Set<string>();
    for (const fFile of footerFiles) {
        const fDoc = parseXmlString(fFile.content.toString());
        const ftrEl = getElementsByTagName(fDoc, "w:ftr")[0];
        if (!ftrEl) continue;
        const text = extractPlainText(ftrEl as Element);
        if (text && !seenFooterTexts.has(text)) {
            seenFooterTexts.add(text);
            content.push({ type: 'footer', text } as OfficeContentNode);
        }
    }

    for (const file of files) {
        if (file.path.match(mediaFileRegex)) continue;
        if (file.path.match(numberingFileRegex)) continue;
        if (file.path.match(relsFileRegex)) continue;
        if (file.path.match(stylesFileRegex)) continue;
        if (file.path.match(footnotesFileRegex)) continue;
        if (file.path.match(endnotesFileRegex)) continue;
        if (file.path.match(commentsFileRegex)) continue;
        if (file.path.match(headerFileRegex)) continue;
        if (file.path.match(footerFileRegex)) continue;

        const documentContent = file.content.toString();
        if (config.includeRawContent) {
            rawContents.push(documentContent);
        }

        const doc = parseXmlString(documentContent);
        const body = getElementsByTagName(doc, "w:body")[0];
        if (body) {
            const bodyChildren = Array.from(body.childNodes);
            for (const child of bodyChildren) {
                if (child.nodeName === 'w:p') {
                    content.push(parseParagraph(child as Element));
                    content.push(...flushPendings());
                } else if (child.nodeName === 'w:tbl') {
                    const { node, sidePending } = parseTable(child as Element);
                    content.push(node, ...sidePending);
                }
            }
        }
    }


    // Extract attachments
    if (config.extractAttachments) {
        for (const media of mediaFiles) {
            const attachment = createAttachment(media.path.split('/').pop() || 'image', media.content);
            attachments.push(attachment);

            if (config.ocr) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        attachment.ocrText = (await performOcr(media.content, config.ocrLanguage)).trim();
                    } catch (e) {
                        logWarning(`OCR failed for ${attachment.name}:`, config, e);
                    }
                }
            }
        }

        for (const a of attachments) {
            attachmentByNameAndType.set(`${a.type}:${a.name}`, a);
        }

        // Assign OCR text to image nodes
        if (config.ocr) {
            const assignOcr = (nodes: OfficeContentNode[]) => {
                for (const node of nodes) {
                    if (node.type === 'image' && 'attachmentName' in (node.metadata || {})) {
                        const meta = node.metadata as ImageMetadata;
                        const attachment = attachmentByNameAndType.get(`image:${meta.attachmentName}`);
                        if (attachment && attachment.ocrText) {
                            node.text = attachment.ocrText;
                            attachment.altText = meta.altText;
                        }
                    }
                    if (node.children) {
                        assignOcr(node.children);
                    }
                }
            };
            assignOcr(content);
        }
    }

    if (config.putNotesAtLast && collectedNotes.length > 0) {
        content.push(...collectedNotes);
    }

    /**
     * Converts a table node to a TableBlock.
     */
    const convertTableToBlock = (tableNode: OfficeContentNode): TableBlock => {
        const rows: Array<{ cols: Array<{ value: string }> }> = [];
        
        if (tableNode.children) {
            for (const rowNode of tableNode.children) {
                if (rowNode.type === 'row' && rowNode.children) {
                    const cols: Array<{ value: string }> = [];
                    
                    for (const cellNode of rowNode.children) {
                        if (cellNode.type === 'cell') {
                            cols.push({ value: cellNode.text || '' });
                        }
                    }
                    
                    if (cols.length > 0) {
                        rows.push({ cols });
                    }
                }
            }
        }
        
        return {
            type: 'table',
            rows
        };
    };

    /**
     * Converts a chart node to a ChartBlock.
     */
    const convertChartToBlock = (chartNode: OfficeContentNode, attachmentMap: Map<string, OfficeAttachment>): ChartBlock | null => {
        if (chartNode.type !== 'chart') return null;
        
        const chartMetadata = chartNode.metadata as ChartMetadata | undefined;
        
        if (chartMetadata?.attachmentName) {
            const attachment = attachmentMap.get(`chart:${chartMetadata.attachmentName}`);
            if (attachment?.chartData) {
                return {
                    type: 'chart',
                    chartData: attachment.chartData,
                    chartType: attachment.chartData.chartType
                };
            }
        }
        
        return null;
    };

    /**
     * Extracts blocks and fullText from content nodes in a single traversal (document order).
     */
    const newline = config.newlineDelimiter ?? '\n';
    const blocks: Block[] = [];

    const traverseBlocksAndText = (node: OfficeContentNode): string => {
        if (node.type === 'header') {
            const headerBlock: HeaderBlock = { type: 'header', text: node.text ?? '' };
            blocks.push(headerBlock);
            return '';
        }
        if (node.type === 'footer') {
            const footerBlock: FooterBlock = { type: 'footer', text: node.text ?? '' };
            blocks.push(footerBlock);
            return '';
        }
        if (node.type === 'comment') {
            const meta = node.metadata as CommentMetadata | undefined;
            const commentBlock: CommentBlock = {
                type: 'comment',
                text: node.text ?? '',
                author: meta?.author,
                date: meta?.date,
                anchorText: meta?.anchorText,
                replies: meta?.replies
            };
            blocks.push(commentBlock);
            return '';
        }
        if (node.type === 'insertion') {
            const meta = node.metadata as TrackedChangeMetadata | undefined;
            const insertionBlock: InsertionBlock = {
                type: 'insertion',
                text: node.text ?? '',
                author: meta?.author,
                date: meta?.date
            };
            blocks.push(insertionBlock);
            return node.text ?? '';
        }
        if (node.type === 'deletion') {
            const meta = node.metadata as TrackedChangeMetadata | undefined;
            const deletionBlock: DeletionBlock = {
                type: 'deletion',
                text: node.text ?? '',
                author: meta?.author,
                date: meta?.date
            };
            blocks.push(deletionBlock);
            return '';
        }
        if (node.type === 'table') {
            const tableBlock = convertTableToBlock(node);
            blocks.push(tableBlock);
            const tableText = tableBlock.rows.map(r => r.cols.map(c => c.value).join('\t')).join(newline);
            return tableText;
        }
        if (node.type === 'chart') {
            const chartBlock = convertChartToBlock(node, attachmentByNameAndType);
            if (chartBlock) blocks.push(chartBlock);
            return '';
        }
        if (node.type === 'image') {
            const imageMetadata = node.metadata as ImageMetadata | undefined;
            const attachmentName = imageMetadata?.attachmentName;
            if (attachmentName) {
                const attachment = attachmentByNameAndType.get(`image:${attachmentName}`);
                if (attachment) {
                    blocks.push({
                        type: 'image',
                        buffer: Buffer.from(attachment.data, 'base64'),
                        mimeType: attachment.mimeType,
                        filename: attachment.name
                    });
                }
            }
            return '';
        }
        if (node.text && node.text.trim() && (node.type === 'text' || node.type === 'paragraph' || node.type === 'heading')) {
            blocks.push({ type: 'text', content: node.text.trim() });
            return node.text.trim();
        }
        if (node.children) {
            const parts = node.children.map(traverseBlocksAndText).filter(t => t !== '');
            const delimiter = !node.children[0]?.children ? '' : newline;
            return parts.join(delimiter);
        }
        return '';
    };

    const fullText = content.map(traverseBlocksAndText).filter(t => t !== '').join(newline);

    /**
     * Extracts images list from attachments.
     */
    const extractImagesList = (attachments: OfficeAttachment[]): Array<{ buffer: Buffer; mimeType: string; filename?: string }> => {
        return attachments
            .filter(att => att.type === 'image')
            .map(att => ({
                buffer: Buffer.from(att.data, 'base64'),
                mimeType: att.mimeType,
                filename: att.name
            }));
    };

    const images = extractImagesList(attachments);


    return {
        type: 'docx',
        metadata: { ...metadata, formatting: docDefaults, styleMap: styleMap },
        content: content,
        attachments: attachments,
        fullText,
        blocks,
        images,
        toText: () => fullText
    };
};

