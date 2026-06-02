"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseWord = void 0;
const xmldom_1 = require("@xmldom/xmldom");
const errorUtils_1 = require("../utils/errorUtils");
const imageUtils_1 = require("../utils/imageUtils");
const ocrUtils_1 = require("../utils/ocrUtils");
const xmlUtils_1 = require("../utils/xmlUtils");
const zipUtils_1 = require("../utils/zipUtils");
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
const parseWord = async (buffer, config) => {
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
    const xmlSerializer = new xmldom_1.XMLSerializer();
    // Pre-compiled regexes for run-property boolean tags (used in hot path)
    const REGEX_W_B = /<w:b(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const REGEX_W_I = /<w:i(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const REGEX_W_STRIKE = /<w:strike(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const REGEX_W_DSTRIKE = /<w:dstrike(?:\s+w:val="([^"]+)")?\s*\/?>/;
    const getBoolValFromRegex = (xmlSnippet, regex) => {
        const match = xmlSnippet.match(regex);
        if (match) {
            const val = match[1];
            if (val === undefined)
                return true;
            return val === '1' || val === 'true' || val === 'on';
        }
        return null;
    };
    // Helper to extract formatting from run properties XML string
    const extractFormattingFromXml = (rPr) => {
        const formatting = {};
        const rPrString = xmlSerializer.serializeToString(rPr);
        const bold = getBoolValFromRegex(rPrString, REGEX_W_B);
        if (bold !== null)
            formatting.bold = bold;
        const italic = getBoolValFromRegex(rPrString, REGEX_W_I);
        if (italic !== null)
            formatting.italic = italic;
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
        if (strike !== null)
            formatting.strikethrough = strike;
        else if (dstrike !== null)
            formatting.strikethrough = dstrike;
        // Font size
        const szMatch = rPrString.match(/<w:sz w:val="(\d+)"/);
        if (szMatch)
            formatting.size = (parseInt(szMatch[1]) / 2).toString() + 'pt';
        // Color
        const colorMatch = rPrString.match(/<w:color w:val="([^"]+)"/);
        if (colorMatch && colorMatch[1] !== 'auto')
            formatting.color = '#' + colorMatch[1];
        // Background color (shading)
        const shdMatch = rPrString.match(/<w:shd[^>]*w:fill="([^"]+)"/);
        if (shdMatch && shdMatch[1] !== 'auto')
            formatting.backgroundColor = '#' + shdMatch[1];
        // Highlight (map to backgroundColor)
        const highlightMatch = rPrString.match(/<w:highlight w:val="([^"]+)"/);
        if (highlightMatch && highlightMatch[1] !== 'none') {
            const colorMap = {
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
        }
        else {
            const hAnsiMatch = rPrString.match(/<w:rFonts[^>]*w:hAnsi="([^"]+)"/);
            if (hAnsiMatch)
                formatting.font = hAnsiMatch[1];
        }
        // Subscript/Superscript
        const vertAlignMatch = rPrString.match(/<w:vertAlign w:val="([^"]+)"/);
        if (vertAlignMatch) {
            if (vertAlignMatch[1] === 'subscript')
                formatting.subscript = true;
            if (vertAlignMatch[1] === 'superscript')
                formatting.superscript = true;
        }
        return formatting;
    };
    const files = await (0, zipUtils_1.extractFiles)(buffer, x => !!x.match(documentFileRegex) ||
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
        (!!config.extractAttachments && !!x.match(mediaFileRegex)));
    let corePropsFile;
    let relsFile;
    let numberingFile;
    let stylesFile;
    let footnotesFile;
    let endnotesFile;
    let commentsFile;
    let commentsExtendedFile;
    const headerFiles = [];
    const footerFiles = [];
    for (const f of files) {
        if (f.path.match(corePropsFileRegex))
            corePropsFile = f;
        else if (f.path.match(relsFileRegex))
            relsFile = f;
        else if (f.path.match(numberingFileRegex))
            numberingFile = f;
        else if (f.path.match(stylesFileRegex))
            stylesFile = f;
        else if (f.path.match(footnotesFileRegex))
            footnotesFile = f;
        else if (f.path.match(endnotesFileRegex))
            endnotesFile = f;
        else if (f.path.match(commentsExtendedFileRegex))
            commentsExtendedFile = f;
        else if (f.path.match(commentsFileRegex))
            commentsFile = f;
        else if (f.path.match(headerFileRegex))
            headerFiles.push(f);
        else if (f.path.match(footerFileRegex))
            footerFiles.push(f);
    }
    // Extract metadata
    const metadata = corePropsFile ? (0, xmlUtils_1.parseOfficeMetadata)(corePropsFile.content.toString()) : {};
    const footnoteMap = new Map();
    const endnoteMap = new Map();
    const collectedNotes = [];
    const attachments = [];
    const attachmentByNameAndType = new Map();
    const mediaFiles = files.filter(f => f.path.match(mediaFileRegex));
    // Extract relationships
    const relsMap = {};
    if (relsFile) {
        const relsXml = (0, xmlUtils_1.parseXmlString)(relsFile.content.toString());
        const relationships = (0, xmlUtils_1.getElementsByTagName)(relsXml, "Relationship");
        for (let i = 0; i < relationships.length; i++) {
            const id = relationships[i].getAttribute("Id");
            const target = relationships[i].getAttribute("Target");
            if (id && target) {
                relsMap[id] = target;
            }
        }
    }
    const numberingMap = {};
    if (numberingFile) {
        const numberingXml = (0, xmlUtils_1.parseXmlString)(numberingFile.content.toString());
        const nums = (0, xmlUtils_1.getElementsByTagName)(numberingXml, "w:num");
        const abstractNums = (0, xmlUtils_1.getElementsByTagName)(numberingXml, "w:abstractNum");
        const abstractNumMap = {};
        for (let i = 0; i < abstractNums.length; i++) {
            const abstractNumId = abstractNums[i].getAttribute("w:abstractNumId");
            if (abstractNumId) {
                abstractNumMap[abstractNumId] = abstractNums[i];
            }
        }
        for (let i = 0; i < nums.length; i++) {
            const numId = nums[i].getAttribute("w:numId");
            const abstractNumIdNode = (0, xmlUtils_1.getElementsByTagName)(nums[i], "w:abstractNumId")[0];
            const abstractNumId = abstractNumIdNode?.getAttribute("w:val");
            if (numId && abstractNumId && abstractNumMap[abstractNumId]) {
                numberingMap[numId] = {};
                const lvls = (0, xmlUtils_1.getElementsByTagName)(abstractNumMap[abstractNumId], "w:lvl");
                for (let j = 0; j < lvls.length; j++) {
                    const ilvl = lvls[j].getAttribute("w:ilvl");
                    const numFmtNode = (0, xmlUtils_1.getElementsByTagName)(lvls[j], "w:numFmt")[0];
                    const lvlTextNode = (0, xmlUtils_1.getElementsByTagName)(lvls[j], "w:lvlText")[0];
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
    const styleMap = {};
    let docDefaults = {};
    let defaultParaStyleId = undefined;
    if (stylesFile) {
        const stylesXml = (0, xmlUtils_1.parseXmlString)(stylesFile.content.toString());
        const styles = (0, xmlUtils_1.getElementsByTagName)(stylesXml, "w:style");
        for (let i = 0; i < styles.length; i++) {
            const styleId = styles[i].getAttribute("w:styleId");
            if (styleId) {
                const rPr = (0, xmlUtils_1.getElementsByTagName)(styles[i], "w:rPr")[0];
                const pPr = (0, xmlUtils_1.getElementsByTagName)(styles[i], "w:pPr")[0];
                const formatting = rPr ? extractFormattingFromXml(rPr) : {};
                let alignment = undefined;
                let backgroundColor = undefined;
                if (pPr) {
                    const jc = (0, xmlUtils_1.getElementsByTagName)(pPr, "w:jc")[0];
                    if (jc) {
                        const val = jc.getAttribute("w:val");
                        if (val === 'left' || val === 'center' || val === 'right' || val === 'justify') {
                            alignment = val;
                        }
                    }
                    const shd = (0, xmlUtils_1.getElementsByTagName)(pPr, "w:shd")[0];
                    if (shd) {
                        const fill = shd.getAttribute("w:fill");
                        if (fill && fill !== 'auto')
                            backgroundColor = '#' + fill;
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
        const docDefaultsNode = (0, xmlUtils_1.getElementsByTagName)(stylesXml, "w:docDefaults")[0];
        if (docDefaultsNode) {
            const rPrDefaultNode = (0, xmlUtils_1.getElementsByTagName)(docDefaultsNode, "w:rPrDefault")[0];
            if (rPrDefaultNode) {
                const rPr = (0, xmlUtils_1.getElementsByTagName)(rPrDefaultNode, "w:rPr")[0];
                if (rPr) {
                    docDefaults = extractFormattingFromXml(rPr);
                }
            }
        }
        if (!defaultParaStyleId && styleMap["Normal"]) {
            defaultParaStyleId = "Normal";
        }
    }
    const structuredCommentMap = new Map();
    // Build paraId -> parentParaId map from commentsExtended.xml as a fallback
    // for documents where w14:paraIdParent is absent on comment paragraphs
    const extendedParentMap = new Map();
    if (commentsExtendedFile && !config.ignoreComments) {
        const extDoc = (0, xmlUtils_1.parseXmlString)(commentsExtendedFile.content.toString());
        // getElementsByTagName matches on the serialised prefix literal, not namespace URI.
        // Documents that bind the same namespace to a different prefix would be missed here.
        const commentExNodes = (0, xmlUtils_1.getElementsByTagName)(extDoc, "w15:commentEx");
        for (const node of commentExNodes) {
            const paraId = node.getAttribute("w15:paraId");
            const parentParaId = node.getAttribute("w15:paraIdParent");
            if (paraId && parentParaId) {
                extendedParentMap.set(paraId, parentParaId);
            }
        }
    }
    if (commentsFile && !config.ignoreComments) {
        const commentsDoc = (0, xmlUtils_1.parseXmlString)(commentsFile.content.toString());
        const commentNodes = (0, xmlUtils_1.getElementsByTagName)(commentsDoc, "w:comment");
        const rawComments = [];
        const paraIdToCommentId = new Map();
        for (const commentNode of commentNodes) {
            const id = commentNode.getAttribute("w:id");
            if (!id)
                continue;
            const author = commentNode.getAttribute("w:author") ?? undefined;
            const date = commentNode.getAttribute("w:date") ?? undefined;
            const pNodes = (0, xmlUtils_1.getElementsByTagName)(commentNode, "w:p");
            const firstPara = pNodes[0];
            // w14:paraId and w14:paraIdParent use the w14 namespace
            const firstParaId = firstPara?.getAttribute("w14:paraId") ?? undefined;
            const parentParaId = firstPara?.getAttribute("w14:paraIdParent")
                || (firstParaId ? extendedParentMap.get(firstParaId) : undefined);
            const text = pNodes
                .map(p => (0, xmlUtils_1.getElementsByTagName)(p, "w:t").map(t => t.textContent ?? '').join(''))
                .filter(t => t)
                .join('\n');
            rawComments.push({ id, author, date, text, firstParaId, parentParaId });
            if (firstParaId)
                paraIdToCommentId.set(firstParaId, id);
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
                }
                else {
                    // Parent not found — treat as top-level
                    structuredCommentMap.set(raw.id, {
                        id: raw.id, author: raw.author, date: raw.date, text: raw.text, replies: []
                    });
                }
            }
        }
    }
    // Tracks which comment IDs were referenced in the current paragraph (populated during parseParagraph)
    let pendingCommentRefs = [];
    // Tracked-change nodes collected during paragraph parsing, emitted as top-level content after the paragraph
    let pendingTrackedChanges = [];
    // Anchor text accumulated between commentRangeStart and commentRangeEnd
    const commentAnchorAccum = new Map();
    // Open comment ranges that span paragraph boundaries (keyed by comment ID)
    const activeCommentRanges = new Set();
    const content = [];
    const rawContents = [];
    const numberingState = {};
    const listCounters = {}; // Track item index per listId/level
    // Helper to parse a paragraph node
    const parseParagraph = (pNode, isNoteContext = false, skipReset = false) => {
        const pXml = xmlSerializer.serializeToString(pNode);
        // Check if it's a list item
        const numPr = (0, xmlUtils_1.getElementsByTagName)(pNode, "w:numPr")[0];
        const isList = !!numPr;
        // Check if it's a heading
        const pPr = (0, xmlUtils_1.getElementsByTagName)(pNode, "w:pPr")[0];
        const pStyle = pPr ? (0, xmlUtils_1.getElementsByTagName)(pPr, "w:pStyle")[0] : null;
        const pStyleVal = pStyle ? pStyle.getAttribute("w:val") : null;
        const isHeading = pStyleVal ? (pStyleVal.startsWith("Heading") || pStyleVal === "Title") : false;
        // Extract Paragraph Style Properties
        const styleProps = pStyleVal && styleMap[pStyleVal] ? styleMap[pStyleVal] : { formatting: {} };
        // Extract Alignment
        let alignment = styleProps.alignment;
        if (pPr) {
            const jc = (0, xmlUtils_1.getElementsByTagName)(pPr, "w:jc")[0];
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
            const shd = (0, xmlUtils_1.getElementsByTagName)(pPr, "w:shd")[0];
            if (shd) {
                const fill = shd.getAttribute("w:fill");
                if (fill && fill !== 'auto') {
                    paraBackgroundColor = '#' + fill;
                }
            }
        }
        // Extract paragraph-level run properties
        let paragraphRunFormatting = { ...styleProps.formatting };
        if (pPr) {
            const pPrRPr = (0, xmlUtils_1.getElementsByTagName)(pPr, "w:rPr")[0];
            if (pPrRPr) {
                const pPrFormatting = extractFormattingFromXml(pPrRPr);
                for (const key in pPrFormatting) {
                    const value = pPrFormatting[key];
                    if (value === false) {
                        delete paragraphRunFormatting[key];
                    }
                    else if (value !== undefined) {
                        paragraphRunFormatting[key] = value;
                    }
                }
            }
        }
        // Extract text and children
        let text = '';
        const children = [];
        // Reset pending state for a fresh top-level paragraph. Skip the reset when recursing
        // into text-box content (skipReset=true) so refs collected before the text-box run are
        // preserved, and skip for note contexts (isNoteContext=true) which have their own domain.
        if (!isNoteContext && !skipReset) {
            pendingCommentRefs = [];
            pendingTrackedChanges = [];
        }
        // Traverse children of paragraph (runs, hyperlinks, etc.)
        const processChildNode = (node) => {
            if (node.nodeName === 'w:commentRangeStart' && !config.ignoreComments) {
                const id = node.getAttribute('w:id');
                if (id) {
                    activeCommentRanges.add(id);
                    if (!commentAnchorAccum.has(id))
                        commentAnchorAccum.set(id, '');
                }
            }
            else if (node.nodeName === 'w:commentRangeEnd' && !config.ignoreComments) {
                const id = node.getAttribute('w:id');
                if (id)
                    activeCommentRanges.delete(id);
            }
            else if (node.nodeName === 'w:ins' && !config.ignoreTrackedChanges) {
                const insNode = node;
                const author = insNode.getAttribute('w:author') ?? undefined;
                const date = insNode.getAttribute('w:date') ?? undefined;
                const runs = (0, xmlUtils_1.getElementsByTagName)(insNode, 'w:t');
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
                            metadata: { author, date }
                        });
                    }
                }
            }
            else if (node.nodeName === 'w:del' && !config.ignoreTrackedChanges) {
                const delNode = node;
                const author = delNode.getAttribute('w:author') ?? undefined;
                const date = delNode.getAttribute('w:date') ?? undefined;
                // Deleted text uses w:delText, not w:t
                const runs = (0, xmlUtils_1.getElementsByTagName)(delNode, 'w:delText');
                const deletedText = runs.map(r => r.textContent ?? '').join('');
                if (deletedText && !isNoteContext) {
                    pendingTrackedChanges.push({
                        type: 'deletion',
                        text: deletedText,
                        metadata: { author, date }
                    });
                }
            }
            else if (node.nodeName === 'w:r') {
                const runNode = node;
                const rPr = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:rPr")[0];
                // Formatting
                let formatting = {};
                // Apply paragraph-level formatting
                for (const key in paragraphRunFormatting) {
                    formatting[key] = paragraphRunFormatting[key];
                }
                // Check for run style
                const rStyle = rPr ? (0, xmlUtils_1.getElementsByTagName)(rPr, "w:rStyle")[0] : null;
                const rStyleVal = rStyle ? rStyle.getAttribute("w:val") : pStyleVal;
                if (rStyleVal && styleMap[rStyleVal]) {
                    for (const key in styleMap[rStyleVal].formatting) {
                        formatting[key] = styleMap[rStyleVal].formatting[key];
                    }
                }
                // Apply direct run properties
                if (rPr) {
                    const directFormatting = extractFormattingFromXml(rPr);
                    for (const key in directFormatting) {
                        const value = directFormatting[key];
                        if (value === false) {
                            delete formatting[key];
                        }
                        else if (value !== undefined) {
                            formatting[key] = value;
                        }
                    }
                }
                // Inherit paragraph background
                if (!formatting.backgroundColor && paraBackgroundColor) {
                    formatting.backgroundColor = paraBackgroundColor;
                }
                // Text content
                const tNodes = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:t");
                for (const tNode of tNodes) {
                    const tContent = tNode.textContent || '';
                    text += tContent;
                    // Accumulate anchor text for any open comment ranges
                    for (const id of activeCommentRanges) {
                        commentAnchorAccum.set(id, (commentAnchorAccum.get(id) ?? '') + tContent);
                    }
                    const textNode = {
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
                    const commentRef = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:commentReference")[0];
                    if (commentRef) {
                        const id = commentRef.getAttribute("w:id");
                        if (id && structuredCommentMap.has(id)) {
                            pendingCommentRefs.push(id);
                        }
                    }
                }
                // Images/Drawings
                if (config.extractAttachments) {
                    const drawings = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:drawing");
                    const picts = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:pict");
                    const allImages = [...drawings, ...picts];
                    for (const imgNode of allImages) {
                        const imgXml = xmlSerializer.serializeToString(imgNode);
                        // Extract Alt Text
                        let altText = '';
                        const docPr = (0, xmlUtils_1.getElementsByTagName)(imgNode, "wp:docPr")[0];
                        if (docPr) {
                            altText = docPr.getAttribute("descr") || docPr.getAttribute("title") || '';
                        }
                        // Extract Relationship ID
                        let rId = '';
                        const blip = (0, xmlUtils_1.getElementsByTagName)(imgNode, "a:blip")[0];
                        if (blip) {
                            rId = blip.getAttribute("r:embed") || '';
                        }
                        else {
                            const imagedata = (0, xmlUtils_1.getElementsByTagName)(imgNode, "v:imagedata")[0];
                            if (imagedata) {
                                rId = imagedata.getAttribute("r:id") || '';
                            }
                        }
                        if (rId && relsMap[rId]) {
                            const target = relsMap[rId];
                            const filename = target.split('/').pop();
                            if (filename) {
                                const imageNode = {
                                    type: 'image',
                                    text: '',
                                    metadata: { attachmentName: filename, altText: altText }
                                };
                                if (config.includeRawContent) {
                                    imageNode.rawContent = imgXml;
                                }
                                children.push(imageNode);
                            }
                        }
                        else {
                            const imageNode = {
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
                const txbxContents = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:txbxContent");
                for (const txbx of txbxContents) {
                    const pNodes = (0, xmlUtils_1.getElementsByTagName)(txbx, "w:p");
                    for (const p of pNodes) {
                        const parsed = parseParagraph(p, false, true);
                        if (parsed.text?.trim())
                            children.push(parsed);
                    }
                }
                // Footnotes/Endnotes inside runs
                if (!config.ignoreNotes) {
                    const footnoteRef = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:footnoteReference")[0];
                    if (footnoteRef) {
                        const id = footnoteRef.getAttribute("w:id");
                        if (id && footnoteMap.has(id)) {
                            const noteNodes = footnoteMap.get(id);
                            const noteNode = {
                                type: 'note',
                                text: noteNodes.map((n) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'footnote', noteId: id }
                            };
                            if (config.putNotesAtLast) {
                                collectedNotes.push(noteNode);
                            }
                            else {
                                children.push(noteNode);
                            }
                        }
                    }
                    const endnoteRef = (0, xmlUtils_1.getElementsByTagName)(runNode, "w:endnoteReference")[0];
                    if (endnoteRef) {
                        const id = endnoteRef.getAttribute("w:id");
                        if (id && endnoteMap.has(id)) {
                            const noteNodes = endnoteMap.get(id);
                            const noteNode = {
                                type: 'note',
                                text: noteNodes.map((n) => n.text).join(' '),
                                children: noteNodes,
                                metadata: { noteType: 'endnote', noteId: id }
                            };
                            if (config.putNotesAtLast) {
                                collectedNotes.push(noteNode);
                            }
                            else {
                                children.push(noteNode);
                            }
                        }
                    }
                }
            }
            else if (node.nodeName === 'w:hyperlink') {
                const hlNode = node;
                const rId = hlNode.getAttribute("r:id");
                const anchor = hlNode.getAttribute("w:anchor");
                let linkMetadata;
                if (anchor) {
                    linkMetadata = { link: '#' + anchor, linkType: 'internal' };
                }
                else if (rId && relsMap[rId]) {
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
            const numIdNode = (0, xmlUtils_1.getElementsByTagName)(numPr, "w:numId")[0];
            const ilvlNode = (0, xmlUtils_1.getElementsByTagName)(numPr, "w:ilvl")[0];
            const numId = numIdNode ? numIdNode.getAttribute("w:val") || '0' : '0';
            const ilvl = ilvlNode ? parseInt(ilvlNode.getAttribute("w:val") || '0') : 0;
            let listType = 'ordered';
            let itemIndex = 0;
            if (numId && numberingMap[numId]) {
                const ilvlStr = ilvl.toString();
                if (!numberingState[numId])
                    numberingState[numId] = {};
                if (!numberingState[numId][ilvlStr])
                    numberingState[numId][ilvlStr] = 0;
                numberingState[numId][ilvlStr]++;
                for (let k = ilvl + 1; k < 10; k++) {
                    if (numberingState[numId][k.toString()])
                        numberingState[numId][k.toString()] = 0;
                }
                const numFmt = numberingMap[numId][ilvlStr]?.numFmt || 'decimal';
                listType = numFmt === 'bullet' ? 'unordered' : 'ordered';
                // Track itemIndex (starts at 0, continues across interruptions for same listId)
                if (!listCounters[numId])
                    listCounters[numId] = {};
                if (listCounters[numId][ilvlStr] === undefined) {
                    listCounters[numId][ilvlStr] = 0;
                }
                else {
                    listCounters[numId][ilvlStr]++;
                }
                itemIndex = listCounters[numId][ilvlStr];
            }
            const listNode = {
                type: 'list',
                text: text,
                children: children,
                metadata: {
                    listType,
                    indentation: ilvl,
                    alignment: (alignment || 'left'),
                    listId: numId,
                    itemIndex: itemIndex,
                    style: pStyleVal
                }
            };
            if (config.includeRawContent)
                listNode.rawContent = pXml;
            return listNode;
        }
        else if (isHeading) {
            const level = pStyleVal ? parseInt(pStyleVal.replace("Heading", "")) || 1 : 1;
            const headingNode = {
                type: 'heading',
                text: text,
                children: children,
                metadata: { level, alignment, style: pStyleVal ?? undefined }
            };
            if (config.includeRawContent)
                headingNode.rawContent = pXml;
            return headingNode;
        }
        else {
            const paraNode = {
                type: 'paragraph',
                text: text,
                children: children,
                metadata: { alignment, style: pStyleVal ?? undefined }
            };
            if (config.includeRawContent)
                paraNode.rawContent = pXml;
            return paraNode;
        }
    };
    // Drain pending comment refs and tracked-change nodes into a flat list and reset both arrays.
    // Called after every parseParagraph and parseTable invocation that may have populated them.
    const flushPendings = () => {
        const flushed = [];
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
                    const commentMeta = {
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
    const parseTable = (tblNode) => {
        const rows = [];
        const sidePending = [];
        // Only get direct child rows, not nested table rows
        const trNodes = (0, xmlUtils_1.getDirectChildren)(tblNode, "w:tr");
        for (let rIndex = 0; rIndex < trNodes.length; rIndex++) {
            const trNode = trNodes[rIndex];
            const cells = [];
            // Only get direct child cells, not nested table cells
            const tcNodes = (0, xmlUtils_1.getDirectChildren)(trNode, "w:tc");
            for (let cIndex = 0; cIndex < tcNodes.length; cIndex++) {
                const tcNode = tcNodes[cIndex];
                const cellChildren = [];
                let cellText = '';
                // Cells contain paragraphs (and other block-level elements)
                const cellContentNodes = Array.from(tcNode.childNodes);
                for (const child of cellContentNodes) {
                    if (child.nodeName === 'w:p') {
                        const pNode = parseParagraph(child);
                        cellChildren.push(pNode);
                        cellText += pNode.text;
                        // Flush any comments/tracked-changes from this cell paragraph
                        sidePending.push(...flushPendings());
                    }
                    else if (child.nodeName === 'w:tbl') {
                        // Nested table — cascade its side-pending nodes up to ours
                        const nested = parseTable(child);
                        cellChildren.push(nested.node);
                        sidePending.push(...nested.sidePending);
                        // Don't add nested table text to cell text - it will be handled recursively
                    }
                }
                const getCellText = (node) => {
                    let text = node.text || '';
                    if (node.children && node.children.length > 0) {
                        const childTexts = node.children.map(getCellText).filter(t => t !== '');
                        if (childTexts.length > 0) {
                            text += (text ? ' ' : '') + childTexts.join(' ');
                        }
                    }
                    return text;
                };
                const cellNode = {
                    type: 'cell',
                    text: cellText,
                    children: cellChildren,
                    metadata: { row: rIndex, col: cIndex }
                };
                cellNode.text = getCellText(cellNode);
                cells.push(cellNode);
            }
            const rowNode = {
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
            const footnotesDoc = (0, xmlUtils_1.parseXmlString)(footnotesFile.content.toString());
            const footnoteNodes = (0, xmlUtils_1.getElementsByTagName)(footnotesDoc, "w:footnote");
            for (const node of footnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0")
                    continue;
                const pNodes = (0, xmlUtils_1.getElementsByTagName)(node, "w:p");
                footnoteMap.set(id, pNodes.map(p => parseParagraph(p, true)));
            }
        }
        if (endnotesFile) {
            const endnotesDoc = (0, xmlUtils_1.parseXmlString)(endnotesFile.content.toString());
            const endnoteNodes = (0, xmlUtils_1.getElementsByTagName)(endnotesDoc, "w:endnote");
            for (const node of endnoteNodes) {
                const id = node.getAttribute("w:id");
                if (!id || id === "-1" || id === "0")
                    continue;
                const pNodes = (0, xmlUtils_1.getElementsByTagName)(node, "w:p");
                endnoteMap.set(id, pNodes.map(p => parseParagraph(p, true)));
            }
        }
    }
    // Extract headers and footers, deduplicated by plain text content, before body.
    const extractPlainText = (el) => (0, xmlUtils_1.getElementsByTagName)(el, "w:t").map(t => t.textContent || '').join('').trim();
    const seenHeaderTexts = new Set();
    for (const hFile of headerFiles) {
        const hDoc = (0, xmlUtils_1.parseXmlString)(hFile.content.toString());
        const hdrEl = (0, xmlUtils_1.getElementsByTagName)(hDoc, "w:hdr")[0];
        if (!hdrEl)
            continue;
        const text = extractPlainText(hdrEl);
        if (text && !seenHeaderTexts.has(text)) {
            seenHeaderTexts.add(text);
            content.push({ type: 'header', text });
        }
    }
    const seenFooterTexts = new Set();
    for (const fFile of footerFiles) {
        const fDoc = (0, xmlUtils_1.parseXmlString)(fFile.content.toString());
        const ftrEl = (0, xmlUtils_1.getElementsByTagName)(fDoc, "w:ftr")[0];
        if (!ftrEl)
            continue;
        const text = extractPlainText(ftrEl);
        if (text && !seenFooterTexts.has(text)) {
            seenFooterTexts.add(text);
            content.push({ type: 'footer', text });
        }
    }
    for (const file of files) {
        if (file.path.match(mediaFileRegex))
            continue;
        if (file.path.match(numberingFileRegex))
            continue;
        if (file.path.match(relsFileRegex))
            continue;
        if (file.path.match(stylesFileRegex))
            continue;
        if (file.path.match(footnotesFileRegex))
            continue;
        if (file.path.match(endnotesFileRegex))
            continue;
        if (file.path.match(commentsFileRegex))
            continue;
        if (file.path.match(headerFileRegex))
            continue;
        if (file.path.match(footerFileRegex))
            continue;
        const documentContent = file.content.toString();
        if (config.includeRawContent) {
            rawContents.push(documentContent);
        }
        const doc = (0, xmlUtils_1.parseXmlString)(documentContent);
        const body = (0, xmlUtils_1.getElementsByTagName)(doc, "w:body")[0];
        if (body) {
            const bodyChildren = Array.from(body.childNodes);
            for (const child of bodyChildren) {
                if (child.nodeName === 'w:p') {
                    content.push(parseParagraph(child));
                    content.push(...flushPendings());
                }
                else if (child.nodeName === 'w:tbl') {
                    const { node, sidePending } = parseTable(child);
                    content.push(node, ...sidePending);
                }
            }
        }
    }
    // Extract attachments
    if (config.extractAttachments) {
        for (const media of mediaFiles) {
            const attachment = (0, imageUtils_1.createAttachment)(media.path.split('/').pop() || 'image', media.content);
            attachments.push(attachment);
            if (config.ocr) {
                if (attachment.mimeType.startsWith('image/')) {
                    try {
                        attachment.ocrText = (await (0, ocrUtils_1.performOcr)(media.content, config.ocrLanguage)).trim();
                    }
                    catch (e) {
                        (0, errorUtils_1.logWarning)(`OCR failed for ${attachment.name}:`, config, e);
                    }
                }
            }
        }
        for (const a of attachments) {
            attachmentByNameAndType.set(`${a.type}:${a.name}`, a);
        }
        // Assign OCR text to image nodes
        if (config.ocr) {
            const assignOcr = (nodes) => {
                for (const node of nodes) {
                    if (node.type === 'image' && 'attachmentName' in (node.metadata || {})) {
                        const meta = node.metadata;
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
    const convertTableToBlock = (tableNode) => {
        const rows = [];
        if (tableNode.children) {
            for (const rowNode of tableNode.children) {
                if (rowNode.type === 'row' && rowNode.children) {
                    const cols = [];
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
    const convertChartToBlock = (chartNode, attachmentMap) => {
        if (chartNode.type !== 'chart')
            return null;
        const chartMetadata = chartNode.metadata;
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
    const blocks = [];
    const traverseBlocksAndText = (node) => {
        if (node.type === 'header') {
            const headerBlock = { type: 'header', text: node.text ?? '' };
            blocks.push(headerBlock);
            return '';
        }
        if (node.type === 'footer') {
            const footerBlock = { type: 'footer', text: node.text ?? '' };
            blocks.push(footerBlock);
            return '';
        }
        if (node.type === 'comment') {
            const meta = node.metadata;
            const commentBlock = {
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
            const meta = node.metadata;
            const insertionBlock = {
                type: 'insertion',
                text: node.text ?? '',
                author: meta?.author,
                date: meta?.date
            };
            blocks.push(insertionBlock);
            return node.text ?? '';
        }
        if (node.type === 'deletion') {
            const meta = node.metadata;
            const deletionBlock = {
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
            if (chartBlock)
                blocks.push(chartBlock);
            return '';
        }
        if (node.type === 'image') {
            const imageMetadata = node.metadata;
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
    const extractImagesList = (attachments) => {
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
exports.parseWord = parseWord;
