"use strict";
/**
 * @file textExtractor.ts
 * Module for parsing OOXML WordprocessingML content and extracting structured text.
 */
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.extractStyleMap = extractStyleMap;
exports.extractParagraph = extractParagraph;
exports.extractText = extractText;
exports.resolveListItems = resolveListItems;
var xmlParser_js_1 = require("../utils/xmlParser.js");
var tableExtractor_js_1 = require("./tableExtractor.js");
var imageExtractor_js_1 = require("./imageExtractor.js");
/**
 * Type guard checking whether a value is a non-null Record object.
 *
 * @param value - Value to check.
 * @returns True if value is an object and not an array.
 */
function isRecord(value) {
    return typeof value === 'object' && value !== null && !Array.isArray(value);
}
/**
 * Recursively discovers all w:txbxContent nodes within a drawing or shape container.
 *
 * @param node - Container node.
 * @returns Array of w:txbxContent record nodes found.
 */
function findTxbxContentNodes(node) {
    var results = [];
    function search(curr) {
        if (!isRecord(curr))
            return;
        for (var _i = 0, _a = Object.entries(curr); _i < _a.length; _i++) {
            var _b = _a[_i], key = _b[0], val = _b[1];
            if (key === 'w:txbxContent' || key === 'txbxContent') {
                if (isRecord(val)) {
                    results.push(val);
                }
                else if (Array.isArray(val)) {
                    for (var _c = 0, val_1 = val; _c < val_1.length; _c++) {
                        var item = val_1[_c];
                        if (isRecord(item))
                            results.push(item);
                    }
                }
            }
            else if (isRecord(val)) {
                search(val);
            }
            else if (Array.isArray(val)) {
                for (var _d = 0, val_2 = val; _d < val_2.length; _d++) {
                    var item = val_2[_d];
                    search(item);
                }
            }
        }
    }
    search(node);
    return results;
}
/**
 * Normalizes a value that may be undefined, a single item, or an array into an array.
 *
 * @param value - Input value.
 * @returns An array of items.
 */
function toArray(value) {
    if (value === undefined || value === null) {
        return [];
    }
    return Array.isArray(value) ? value : [value];
}
/**
 * Retrieves a property from a record matching one of several candidate keys.
 *
 * @param record - Target record.
 * @param keys - Candidate key names.
 * @returns The matching property value, or undefined.
 */
function getProperty(record) {
    var keys = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        keys[_i - 1] = arguments[_i];
    }
    for (var _a = 0, keys_1 = keys; _a < keys_1.length; _a++) {
        var key = keys_1[_a];
        if (key in record && record[key] !== undefined) {
            return record[key];
        }
    }
    return undefined;
}
/**
 * Retrieves a numeric attribute from a record matching one of several candidate keys.
 * Tries each key, parses as integer, and returns the first valid number or undefined.
 *
 * @param record - Target record.
 * @param keys - Candidate attribute keys.
 * @returns The parsed integer value, or undefined if none found or not a number.
 */
function getNumericAttr(record) {
    var keys = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        keys[_i - 1] = arguments[_i];
    }
    for (var _a = 0, keys_2 = keys; _a < keys_2.length; _a++) {
        var key = keys_2[_a];
        if (key in record && record[key] !== undefined && record[key] !== null) {
            var val = parseInt(String(record[key]), 10);
            if (!isNaN(val)) {
                return val;
            }
        }
    }
    return undefined;
}
/**
 * Checks whether an OOXML toggle property (such as w:b or w:i) represents an active state.
 * In OOXML, `<w:b/>`, `<w:b w:val="true"/>`, `<w:b w:val="1"/>`, and `<w:b w:val="on"/>` are active.
 *
 * @param prop - The property node from parsed XML.
 * @returns True if active, false otherwise.
 */
function isToggleActive(prop) {
    if (prop === undefined || prop === null) {
        return false;
    }
    if (typeof prop === 'boolean') {
        return prop;
    }
    if (typeof prop === 'number') {
        return prop !== 0;
    }
    if (typeof prop === 'string') {
        var s = prop.toLowerCase().trim();
        return s !== 'false' && s !== '0' && s !== 'off';
    }
    if (isRecord(prop)) {
        var val = getProperty(prop, '@_w:val', '@_val');
        if (val === undefined) {
            // Element present without attributes, e.g. <w:b/> -> defaults to true
            return true;
        }
        if (typeof val === 'boolean') {
            return val;
        }
        var valStr = String(val).toLowerCase().trim();
        return valStr !== 'false' && valStr !== '0' && valStr !== 'off' && valStr !== 'none';
    }
    return false;
}
/**
 * Checks whether an underline property (w:u) represents active underlining.
 *
 * @param prop - The property node from parsed XML.
 * @returns True if active, false otherwise.
 */
function isUnderlineActive(prop) {
    if (prop === undefined || prop === null) {
        return false;
    }
    if (typeof prop === 'boolean') {
        return prop;
    }
    if (typeof prop === 'string') {
        var s = prop.toLowerCase().trim();
        return s !== 'none' && s !== 'false' && s !== '0';
    }
    if (isRecord(prop)) {
        var val = getProperty(prop, '@_w:val', '@_val');
        if (val === undefined) {
            return true;
        }
        var valStr = String(val).toLowerCase().trim();
        return valStr !== 'none' && valStr !== 'false' && valStr !== '0';
    }
    return false;
}
/**
 * Extracts raw string content from an OOXML text node (w:t), taking into account
 * attributes such as xml:space="preserve".
 *
 * @param tNode - The w:t node in parsed XML.
 * @returns The extracted string value.
 */
function extractTextFromTNode(tNode) {
    if (tNode === undefined || tNode === null) {
        return '';
    }
    if (typeof tNode === 'string') {
        return tNode;
    }
    if (typeof tNode === 'number' || typeof tNode === 'boolean') {
        return String(tNode);
    }
    if (isRecord(tNode)) {
        var textVal = tNode['#text'];
        if (textVal !== undefined && textVal !== null) {
            return String(textVal);
        }
        return '';
    }
    return '';
}
/**
 * Extracts a single TextRun from a parsed w:r node.
 * If images are present in w:drawing and imageMap is provided, extracts DocxImage metadata.
 *
 * @param rNode - Parsed w:r node.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @param paragraphImages - Optional array accumulating images in the paragraph.
 * @returns Extracted TextRun, or null if no valid run could be parsed.
 */
function extractRun(rNode, imageMap, paragraphImages, defaultColor) {
    if (!isRecord(rNode)) {
        return null;
    }
    // Extract run properties (w:rPr / rPr)
    var rPrRaw = getProperty(rNode, 'w:rPr', 'rPr');
    var rPr = isRecord(rPrRaw) ? rPrRaw : undefined;
    var bold;
    var italic;
    var underline;
    var color;
    if (rPr) {
        var bNode = getProperty(rPr, 'w:b', 'b');
        if (isToggleActive(bNode)) {
            bold = true;
        }
        var iNode = getProperty(rPr, 'w:i', 'i');
        if (isToggleActive(iNode)) {
            italic = true;
        }
        var uNode = getProperty(rPr, 'w:u', 'u');
        if (isUnderlineActive(uNode)) {
            underline = true;
        }
        var colorNode = getProperty(rPr, 'w:color', 'color');
        if (isRecord(colorNode)) {
            var colorVal = getProperty(colorNode, '@_w:val', '@_val', '@w:val', 'val');
            if (colorVal && typeof colorVal === 'string' && colorVal !== 'auto' && colorVal !== '000000') {
                color = colorVal;
            }
        }
        else if (defaultColor) {
            color = defaultColor;
        }
    }
    else if (defaultColor) {
        color = defaultColor;
    }
    // Extract text nodes (w:t)
    var tNodes = toArray(getProperty(rNode, 'w:t', 't'));
    var text = '';
    for (var _i = 0, tNodes_1 = tNodes; _i < tNodes_1.length; _i++) {
        var tNode = tNodes_1[_i];
        text += extractTextFromTNode(tNode);
    }
    // Handle special inline elements such as w:tab and w:br
    if (getProperty(rNode, 'w:tab', 'tab') !== undefined) {
        text += '\t';
    }
    var pageBreak;
    var brNodes = toArray(getProperty(rNode, 'w:br', 'br'));
    for (var _a = 0, brNodes_1 = brNodes; _a < brNodes_1.length; _a++) {
        var br = brNodes_1[_a];
        if (isRecord(br)) {
            var brType = getProperty(br, '@_w:type', '@_type', '@w:type', 'type');
            if (brType === 'page') {
                pageBreak = true;
            }
            else {
                text += '\n';
            }
        }
        else if (br !== undefined && br !== null) {
            text += '\n';
        }
    }
    // Check for w:drawing in the run node
    var drawingNode = getProperty(rNode, 'w:drawing', 'drawing');
    if (drawingNode !== undefined && imageMap && paragraphImages && isRecord(rNode)) {
        var images = (0, imageExtractor_js_1.extractImageMetadata)(rNode, imageMap);
        if (images.length > 0) {
            paragraphImages.push.apply(paragraphImages, images);
        }
    }
    // Return run even if text is empty, as long as it was a valid run element
    var run = __assign(__assign(__assign(__assign(__assign({ text: text }, (bold ? { bold: true } : {})), (italic ? { italic: true } : {})), (underline ? { underline: true } : {})), (color ? { color: color } : {})), (pageBreak ? { pageBreak: true } : {}));
    return run;
}
/**
 * Recursively extracts all run nodes (w:r) from a paragraph or container,
 * including runs nested inside hyperlinks (w:hyperlink).
 *
 * @param container - Parsed paragraph or child container node.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @param paragraphImages - Optional array accumulating images in the paragraph.
 * @param defaultColor - Optional default text color inherited from paragraph properties.
 * @returns Array of parsed TextRun objects.
 */
function extractRunsFromContainer(container, imageMap, paragraphImages, defaultColor) {
    var runs = [];
    // Direct runs (w:r)
    var directRuns = toArray(getProperty(container, 'w:r', 'r'));
    for (var _i = 0, directRuns_1 = directRuns; _i < directRuns_1.length; _i++) {
        var rNode = directRuns_1[_i];
        var run = extractRun(rNode, imageMap, paragraphImages, defaultColor);
        if (run !== null) {
            runs.push(run);
        }
        // Check for text boxes in drawings or pict elements inside this run
        if (isRecord(rNode)) {
            var drawingNodes = toArray(getProperty(rNode, 'w:drawing', 'drawing', 'w:pict', 'pict'));
            for (var _a = 0, drawingNodes_1 = drawingNodes; _a < drawingNodes_1.length; _a++) {
                var drawing = drawingNodes_1[_a];
                var txbxNodes = findTxbxContentNodes(drawing);
                for (var _b = 0, txbxNodes_1 = txbxNodes; _b < txbxNodes_1.length; _b++) {
                    var txbx = txbxNodes_1[_b];
                    var pNodes = toArray(getProperty(txbx, 'w:p', 'p'));
                    for (var pIdx = 0; pIdx < pNodes.length; pIdx++) {
                        var p = pNodes[pIdx];
                        if (isRecord(p)) {
                            if (pIdx > 0 && runs.length > 0) {
                                runs.push({ text: '\n' });
                            }
                            runs.push.apply(runs, extractRunsFromContainer(p, imageMap, paragraphImages));
                        }
                    }
                }
            }
        }
    }
    // Runs nested inside hyperlinks (w:hyperlink)
    var hyperlinks = toArray(getProperty(container, 'w:hyperlink', 'hyperlink'));
    for (var _c = 0, hyperlinks_1 = hyperlinks; _c < hyperlinks_1.length; _c++) {
        var hyperlink = hyperlinks_1[_c];
        if (isRecord(hyperlink)) {
            runs.push.apply(runs, extractRunsFromContainer(hyperlink, imageMap, paragraphImages, defaultColor));
        }
    }
    // Also check container directly for drawing/pict nodes containing text boxes
    var directDrawings = toArray(getProperty(container, 'w:drawing', 'drawing', 'w:pict', 'pict'));
    for (var _d = 0, directDrawings_1 = directDrawings; _d < directDrawings_1.length; _d++) {
        var drawing = directDrawings_1[_d];
        var txbxNodes = findTxbxContentNodes(drawing);
        for (var _e = 0, txbxNodes_2 = txbxNodes; _e < txbxNodes_2.length; _e++) {
            var txbx = txbxNodes_2[_e];
            var pNodes = toArray(getProperty(txbx, 'w:p', 'p'));
            for (var pIdx = 0; pIdx < pNodes.length; pIdx++) {
                var p = pNodes[pIdx];
                if (isRecord(p)) {
                    if (pIdx > 0 && runs.length > 0) {
                        runs.push({ text: '\n' });
                    }
                    runs.push.apply(runs, extractRunsFromContainer(p, imageMap, paragraphImages));
                }
            }
        }
    }
    return runs;
}
/**
 * Extracts paragraph styles and their formatting properties (including alignment) from word/styles.xml.
 * Resolves basedOn inheritance chains so derived styles inherit parent alignment.
 *
 * @param stylesXml - Raw XML string from `word/styles.xml`.
 * @returns Map of style ID (and style name) to DocxStyleInfo.
 */
function extractStyleMap(stylesXml) {
    var map = new Map();
    if (typeof stylesXml !== 'string' || stylesXml.trim().length === 0) {
        return map;
    }
    var parsed;
    try {
        parsed = (0, xmlParser_js_1.parseXml)(stylesXml);
    }
    catch (_a) {
        return map;
    }
    var stylesRoot = getProperty(parsed, 'w:styles', 'styles');
    var rootObj = isRecord(stylesRoot) ? stylesRoot : parsed;
    var styleNodes = toArray(getProperty(rootObj, 'w:style', 'style'));
    for (var _i = 0, styleNodes_1 = styleNodes; _i < styleNodes_1.length; _i++) {
        var styleNode = styleNodes_1[_i];
        if (!isRecord(styleNode)) {
            continue;
        }
        var type = getProperty(styleNode, '@_w:type', '@_type', '@w:type', 'type');
        var typeStr = type !== undefined ? String(type).toLowerCase().trim() : undefined;
        if (typeStr !== undefined && typeStr !== 'paragraph' && typeStr !== 'table') {
            continue;
        }
        var styleId = getProperty(styleNode, '@_w:styleId', '@_styleId', '@w:styleId', 'styleId');
        if (styleId === undefined || styleId === null) {
            continue;
        }
        var styleIdStr = String(styleId);
        var name_1 = void 0;
        var nameNode = getProperty(styleNode, 'w:name', 'name');
        if (isRecord(nameNode)) {
            var val = getProperty(nameNode, '@_w:val', '@_val', '@w:val', 'val');
            if (val !== undefined && val !== null) {
                name_1 = String(val);
            }
        }
        var basedOn = void 0;
        var basedOnNode = getProperty(styleNode, 'w:basedOn', 'basedOn');
        if (isRecord(basedOnNode)) {
            var val = getProperty(basedOnNode, '@_w:val', '@_val', '@w:val', 'val');
            if (val !== undefined && val !== null) {
                basedOn = String(val);
            }
        }
        var isDefault = void 0;
        var defaultVal = getProperty(styleNode, '@_w:default', '@_default', '@w:default', 'default');
        if (defaultVal === '1' || defaultVal === 1 || defaultVal === true || defaultVal === 'true') {
            isDefault = true;
        }
        var align = void 0;
        var indent = void 0;
        var spacing = void 0;
        var pPr = getProperty(styleNode, 'w:pPr', 'pPr');
        if (isRecord(pPr)) {
            var jcRaw = getProperty(pPr, 'w:jc', 'jc');
            var jcVal = void 0;
            if (isRecord(jcRaw)) {
                var val = getProperty(jcRaw, '@_w:val', '@_val', '@w:val', 'val');
                if (val !== undefined && val !== null) {
                    jcVal = String(val).toLowerCase().trim();
                }
            }
            else if (typeof jcRaw === 'string') {
                jcVal = jcRaw.toLowerCase().trim();
            }
            if (jcVal === 'center') {
                align = 'center';
            }
            else if (jcVal === 'right') {
                align = 'right';
            }
            else if (jcVal === 'both' || jcVal === 'justify') {
                align = 'justify';
            }
            else if (jcVal === 'left') {
                align = 'left';
            }
            var indNode = getProperty(pPr, 'w:ind', 'ind');
            if (isRecord(indNode)) {
                var left = getNumericAttr(indNode, '@_w:left', '@_left', '@w:left', 'left', '@_w:start', '@w:start');
                var right = getNumericAttr(indNode, '@_w:right', '@_right', '@w:right', 'right', '@_w:end', '@w:end');
                var firstLine = getNumericAttr(indNode, '@_w:firstLine', '@_firstLine', '@w:firstLine', 'firstLine');
                var hanging = getNumericAttr(indNode, '@_w:hanging', '@_hanging', '@w:hanging', 'hanging');
                if ((left !== undefined && left > 0) ||
                    (right !== undefined && right > 0) ||
                    (firstLine !== undefined && firstLine > 0) ||
                    (hanging !== undefined && hanging > 0)) {
                    indent = {};
                    if (left !== undefined && left > 0)
                        indent.left = left;
                    if (right !== undefined && right > 0)
                        indent.right = right;
                    if (firstLine !== undefined && firstLine > 0)
                        indent.firstLine = firstLine;
                    if (hanging !== undefined && hanging > 0)
                        indent.hanging = hanging;
                }
            }
            var spacingNode = getProperty(pPr, 'w:spacing', 'spacing');
            if (isRecord(spacingNode)) {
                var after = getNumericAttr(spacingNode, '@_w:after', '@_after', '@w:after', 'after');
                var before = getNumericAttr(spacingNode, '@_w:before', '@_before', '@w:before', 'before');
                var line = getNumericAttr(spacingNode, '@_w:line', '@_line', '@w:line', 'line');
                if (after !== undefined || before !== undefined || line !== undefined) {
                    spacing = {};
                    if (after !== undefined)
                        spacing.after = after;
                    if (before !== undefined)
                        spacing.before = before;
                    if (line !== undefined)
                        spacing.line = line;
                }
            }
        }
        var info = __assign(__assign(__assign(__assign(__assign(__assign({ styleId: styleIdStr }, (name_1 !== undefined ? { name: name_1 } : {})), (basedOn !== undefined ? { basedOn: basedOn } : {})), (align !== undefined ? { align: align } : {})), (indent !== undefined ? { indent: indent } : {})), (spacing !== undefined ? { spacing: spacing } : {})), (isDefault ? { isDefault: true } : {}));
        map.set(styleIdStr, info);
        if (name_1 && name_1 !== styleIdStr && !map.has(name_1)) {
            map.set(name_1, info);
        }
    }
    // Resolve basedOn inheritance chains for align, indent, and spacing
    for (var _b = 0, _c = map.values(); _b < _c.length; _b++) {
        var info = _c[_b];
        if ((!info.align || !info.indent || !info.spacing) && info.basedOn) {
            var currentParentId = info.basedOn;
            var visited = new Set([info.styleId]);
            while (currentParentId && !visited.has(currentParentId)) {
                visited.add(currentParentId);
                var parent_1 = map.get(currentParentId);
                if (!parent_1)
                    break;
                if (!info.align && parent_1.align) {
                    info.align = parent_1.align;
                }
                if (!info.indent && parent_1.indent) {
                    info.indent = __assign({}, parent_1.indent);
                }
                if (!info.spacing && parent_1.spacing) {
                    info.spacing = __assign({}, parent_1.spacing);
                }
                if (info.align && info.indent && info.spacing)
                    break;
                currentParentId = parent_1.basedOn;
            }
        }
    }
    return map;
}
/**
 * Extracts a Paragraph object from a parsed w:p node.
 *
 * @param pNode - Parsed w:p XML node.
 * @param options - Parser options.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @param styleMap - Optional mapping of style identifiers to DocxStyleInfo for style inheritance.
 * @returns Extracted Paragraph.
 */
function extractParagraph(pNode, options, imageMap, styleMap) {
    if (!isRecord(pNode)) {
        return {
            text: '',
            runs: [],
            isEmpty: true,
        };
    }
    // Extract paragraph properties (w:pPr / pPr)
    var pPrRaw = getProperty(pNode, 'w:pPr', 'pPr');
    var pPr = isRecord(pPrRaw) ? pPrRaw : undefined;
    var style;
    if ((options === null || options === void 0 ? void 0 : options.includeStyles) !== false && pPr) {
        var pStyleRaw = getProperty(pPr, 'w:pStyle', 'pStyle');
        if (isRecord(pStyleRaw)) {
            var styleVal = getProperty(pStyleRaw, '@_w:val', '@_val');
            if (styleVal !== undefined && styleVal !== null) {
                style = String(styleVal);
            }
        }
    }
    // Check for section break properties (w:sectPr / sectPr)
    var sectionBreak;
    if (pPr) {
        var sectPrRaw = getProperty(pPr, 'w:sectPr', 'sectPr');
        if (isRecord(sectPrRaw)) {
            var columnCount = 1;
            var colsRaw = getProperty(sectPrRaw, 'w:cols', 'cols');
            if (isRecord(colsRaw)) {
                var numVal = getProperty(colsRaw, '@_w:num', '@_num', '@w:num', 'num', '@_w:val', '@_val', 'val');
                if (numVal !== undefined && numVal !== null) {
                    var parsed = parseInt(String(numVal), 10);
                    if (!isNaN(parsed)) {
                        columnCount = parsed;
                    }
                }
            }
            var sectPageBreak = false;
            var typeRaw = getProperty(sectPrRaw, 'w:type', 'type');
            var typeVal = void 0;
            if (isRecord(typeRaw)) {
                var val = getProperty(typeRaw, '@_w:val', '@_val', '@w:val', 'val');
                if (val !== undefined && val !== null) {
                    typeVal = String(val).toLowerCase().trim();
                }
            }
            else if (typeof typeRaw === 'string') {
                typeVal = typeRaw.toLowerCase().trim();
            }
            if (!typeVal || typeVal === 'nextpage' || typeVal === 'oddpage' || typeVal === 'evenpage') {
                sectPageBreak = true;
            }
            if (columnCount > 1 || sectPageBreak || isRecord(colsRaw)) {
                sectionBreak = __assign({ columnCount: columnCount }, (sectPageBreak ? { pageBreak: true } : {}));
            }
            else if (isRecord(sectPrRaw)) {
                // Even a 1-col section break (no cols element) is meaningful as a boundary marker
                sectionBreak = { columnCount: 1 };
            }
        }
    }
    // Check for paragraph default color in w:pPr > w:rPr > w:color
    var defaultColor;
    if (pPr) {
        var pRprRaw = getProperty(pPr, 'w:rPr', 'rPr');
        if (isRecord(pRprRaw)) {
            var colorNode = getProperty(pRprRaw, 'w:color', 'color');
            if (isRecord(colorNode)) {
                var colorVal = getProperty(colorNode, '@_w:val', '@_val', '@w:val', 'val');
                if (colorVal && typeof colorVal === 'string' && colorVal !== 'auto' && colorVal !== '000000') {
                    defaultColor = colorVal;
                }
            }
        }
    }
    // Check for w:pageBreakBefore in pPr
    var pageBreakBefore;
    if (pPr) {
        var pbbNode = getProperty(pPr, 'w:pageBreakBefore', 'pageBreakBefore');
        if (pbbNode !== undefined && pbbNode !== null) {
            if (isRecord(pbbNode)) {
                var val = getProperty(pbbNode, '@_w:val', '@_val', '@w:val', 'val');
                if (val !== undefined && val !== null) {
                    var valStr = String(val).toLowerCase().trim();
                    if (valStr !== '0' && valStr !== 'false' && valStr !== 'off') {
                        pageBreakBefore = true;
                    }
                }
                else {
                    pageBreakBefore = true;
                }
            }
            else if (typeof pbbNode === 'boolean') {
                if (pbbNode)
                    pageBreakBefore = true;
            }
            else if (typeof pbbNode === 'number') {
                if (pbbNode !== 0)
                    pageBreakBefore = true;
            }
            else if (typeof pbbNode === 'string') {
                var s = pbbNode.toLowerCase().trim();
                if (s !== '0' && s !== 'false' && s !== 'off') {
                    pageBreakBefore = true;
                }
            }
            else {
                pageBreakBefore = true;
            }
        }
    }
    // Check for list item numbering (w:numPr / numPr)
    var listItem;
    if (pPr) {
        var numPrRaw = getProperty(pPr, 'w:numPr', 'numPr');
        if (isRecord(numPrRaw)) {
            var numId = 0;
            var numIdNode = getProperty(numPrRaw, 'w:numId', 'numId');
            if (isRecord(numIdNode)) {
                var val = getProperty(numIdNode, '@_w:val', '@_val', '@w:val', 'val');
                if (val !== undefined && val !== null) {
                    var parsed = parseInt(String(val), 10);
                    if (!isNaN(parsed)) {
                        numId = parsed;
                    }
                }
            }
            else if (typeof numIdNode === 'number') {
                numId = Math.trunc(numIdNode);
            }
            else if (typeof numIdNode === 'string') {
                var parsed = parseInt(numIdNode, 10);
                if (!isNaN(parsed)) {
                    numId = parsed;
                }
            }
            // If numId is not 0, it is an active list item
            if (numId !== 0) {
                var level = 0;
                var ilvlNode = getProperty(numPrRaw, 'w:ilvl', 'ilvl');
                if (isRecord(ilvlNode)) {
                    var val = getProperty(ilvlNode, '@_w:val', '@_val', '@w:val', 'val');
                    if (val !== undefined && val !== null) {
                        var parsed = parseInt(String(val), 10);
                        if (!isNaN(parsed)) {
                            level = parsed;
                        }
                    }
                }
                else if (typeof ilvlNode === 'number') {
                    level = Math.trunc(ilvlNode);
                }
                else if (typeof ilvlNode === 'string') {
                    var parsed = parseInt(ilvlNode, 10);
                    if (!isNaN(parsed)) {
                        level = parsed;
                    }
                }
                listItem = {
                    numId: numId,
                    level: level,
                    abstractNumId: 0,
                    listType: 'bullet',
                };
            }
        }
    }
    // Check for paragraph alignment (w:jc / jc)
    var align;
    if (pPr) {
        var jcRaw = getProperty(pPr, 'w:jc', 'jc');
        var jcVal = void 0;
        if (isRecord(jcRaw)) {
            var val = getProperty(jcRaw, '@_w:val', '@_val', '@w:val', 'val');
            if (val !== undefined && val !== null) {
                jcVal = String(val).toLowerCase().trim();
            }
        }
        else if (typeof jcRaw === 'string') {
            jcVal = jcRaw.toLowerCase().trim();
        }
        if (jcVal === 'center') {
            align = 'center';
        }
        else if (jcVal === 'right') {
            align = 'right';
        }
        else if (jcVal === 'both' || jcVal === 'justify') {
            align = 'justify';
        }
    }
    // Inherit alignment from style if not explicitly set on paragraph
    if (align === undefined && style !== undefined && styleMap) {
        var styleInfo = styleMap.get(style);
        if (styleInfo === null || styleInfo === void 0 ? void 0 : styleInfo.align) {
            align = styleInfo.align;
        }
    }
    // Check for paragraph indentation (w:ind / ind)
    var indent;
    if (pPr) {
        var indNode = getProperty(pPr, 'w:ind', 'ind');
        if (isRecord(indNode)) {
            var left = getNumericAttr(indNode, '@_w:left', '@_left', '@w:left', 'left', '@_w:start', '@w:start');
            var right = getNumericAttr(indNode, '@_w:right', '@_right', '@w:right', 'right', '@_w:end', '@w:end');
            var firstLine = getNumericAttr(indNode, '@_w:firstLine', '@_firstLine', '@w:firstLine', 'firstLine');
            var hanging = getNumericAttr(indNode, '@_w:hanging', '@_hanging', '@w:hanging', 'hanging');
            // Only store if any value is non-zero
            if ((left !== undefined && left > 0) ||
                (right !== undefined && right > 0) ||
                (firstLine !== undefined && firstLine > 0) ||
                (hanging !== undefined && hanging > 0)) {
                indent = {};
                if (left !== undefined && left > 0)
                    indent.left = left;
                if (right !== undefined && right > 0)
                    indent.right = right;
                if (firstLine !== undefined && firstLine > 0)
                    indent.firstLine = firstLine;
                if (hanging !== undefined && hanging > 0)
                    indent.hanging = hanging;
            }
        }
    }
    // Inherit indentation from style if not explicitly set on paragraph
    if (indent === undefined && style !== undefined && styleMap) {
        var styleInfo = styleMap.get(style);
        if (styleInfo === null || styleInfo === void 0 ? void 0 : styleInfo.indent) {
            indent = __assign({}, styleInfo.indent);
        }
    }
    // Check for paragraph spacing (w:spacing / spacing)
    var spacing;
    if (pPr) {
        var spacingNode = getProperty(pPr, 'w:spacing', 'spacing');
        if (isRecord(spacingNode)) {
            var after = getNumericAttr(spacingNode, '@_w:after', '@_after', '@w:after', 'after');
            var before = getNumericAttr(spacingNode, '@_w:before', '@_before', '@w:before', 'before');
            var line = getNumericAttr(spacingNode, '@_w:line', '@_line', '@w:line', 'line');
            if (after !== undefined || before !== undefined || line !== undefined) {
                spacing = {};
                if (after !== undefined)
                    spacing.after = after;
                if (before !== undefined)
                    spacing.before = before;
                if (line !== undefined)
                    spacing.line = line;
            }
        }
    }
    // Inherit spacing from style if not explicitly set on paragraph
    if (spacing === undefined && style !== undefined && styleMap) {
        var styleInfo = styleMap.get(style);
        if (styleInfo === null || styleInfo === void 0 ? void 0 : styleInfo.spacing) {
            spacing = __assign({}, styleInfo.spacing);
        }
    }
    // Check for paragraph borders (w:pBdr / pBdr)
    var hasBorderTop;
    var hasBorderBottom;
    if (pPr) {
        var pBdrRaw = getProperty(pPr, 'w:pBdr', 'pBdr');
        if (isRecord(pBdrRaw)) {
            var bottomBdr = getProperty(pBdrRaw, 'w:bottom', 'bottom');
            if (isRecord(bottomBdr)) {
                var val = getProperty(bottomBdr, '@_w:val', '@_val', '@w:val', 'val');
                if (typeof val === 'string' && val.toLowerCase() !== 'none' && val.toLowerCase() !== 'nil') {
                    hasBorderBottom = true;
                }
            }
            var topBdr = getProperty(pBdrRaw, 'w:top', 'top');
            if (isRecord(topBdr)) {
                var val = getProperty(topBdr, '@_w:val', '@_val', '@w:val', 'val');
                if (typeof val === 'string' && val.toLowerCase() !== 'none' && val.toLowerCase() !== 'nil') {
                    hasBorderTop = true;
                }
            }
        }
    }
    var paragraphImages = [];
    var runs = extractRunsFromContainer(pNode, imageMap, paragraphImages, defaultColor);
    if (defaultColor) {
        for (var _i = 0, runs_1 = runs; _i < runs_1.length; _i++) {
            var run = runs_1[_i];
            if (!run.color && run.text && run.text.trim().length > 0) {
                run.color = defaultColor;
            }
        }
    }
    var text = runs.map(function (r) { return r.text; }).join('');
    var hasPageBreakRun = runs.some(function (r) { return r.pageBreak; });
    var isEmpty = (runs.length === 0 || text.length === 0) &&
        paragraphImages.length === 0 &&
        !hasPageBreakRun;
    return __assign(__assign(__assign(__assign(__assign(__assign(__assign(__assign(__assign(__assign(__assign({ text: text, runs: runs }, (style !== undefined ? { style: style } : {})), (align !== undefined ? { align: align } : {})), (indent !== undefined ? { indent: indent } : {})), (spacing !== undefined ? { spacing: spacing } : {})), (hasBorderTop ? { hasBorderTop: true } : {})), (hasBorderBottom ? { hasBorderBottom: true } : {})), (pageBreakBefore ? { pageBreakBefore: true } : {})), (isEmpty ? { isEmpty: true } : {})), (listItem !== undefined ? { listItem: listItem } : {})), (paragraphImages.length > 0 ? { images: paragraphImages } : {})), (sectionBreak !== undefined ? { sectionBreak: sectionBreak } : {}));
}
/**
 * Parses raw WordprocessingML XML content into a structured DocxDocument.
 *
 * Navigates the OOXML hierarchy: `w:document` -> `w:body` -> `w:p` / `w:tbl` / `w:sdt`.
 * Handles formatting indicators (`w:b`, `w:i`, `w:u`), paragraph styles (`w:pStyle`),
 * multi-column sections, empty paragraphs, whitespace preservation, text boxes, and images.
 * Preserves document order via bodyItems.
 *
 * @param rawXml - Raw XML string from `word/document.xml`.
 * @param options - Optional parser configuration.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @returns The structured DocxDocument containing paragraphs, plain text, and body items.
 * @throws {DocxParseError} If the XML is malformed or cannot be parsed.
 */
function extractText(rawXml, options, imageMap) {
    var _a;
    var parsed = (0, xmlParser_js_1.parseXml)(rawXml);
    // Navigate to root document element
    var docRoot = getProperty(parsed, 'w:document', 'document');
    var rootObj = isRecord(docRoot) ? docRoot : parsed;
    // Navigate to body element
    var bodyRoot = getProperty(rootObj, 'w:body', 'body');
    var bodyObj = isRecord(bodyRoot) ? bodyRoot : rootObj;
    // Check for w:sectPr directly in bodyObj (final section properties)
    var sections = [];
    if (isRecord(bodyObj)) {
        var bodySectPr = getProperty(bodyObj, 'w:sectPr', 'sectPr');
        if (isRecord(bodySectPr)) {
            var columnCount = 1;
            var colsNode = getProperty(bodySectPr, 'w:cols', 'cols');
            if (isRecord(colsNode)) {
                var numVal = getProperty(colsNode, '@_w:num', '@_num', '@w:num', 'num', '@_w:val', '@_val', 'val');
                if (numVal !== undefined && numVal !== null) {
                    var parsed_1 = parseInt(String(numVal), 10);
                    if (!isNaN(parsed_1)) {
                        columnCount = parsed_1;
                    }
                }
            }
            if (columnCount > 1) {
                sections.push({ columnCount: columnCount });
            }
        }
    }
    var pNodes = toArray(getProperty(bodyObj, 'w:p', 'p'));
    var tblNodes = toArray(getProperty(bodyObj, 'w:tbl', 'tbl'));
    var sdtNodes = toArray(getProperty(bodyObj, 'w:sdt', 'sdt'));
    var styleMap = (options === null || options === void 0 ? void 0 : options.stylesXml) ? extractStyleMap(options.stylesXml) : undefined;
    var orderedParagraphs = [];
    var orderedTables = [];
    var bodyItems = [];
    var preserveOrderSuccess = false;
    try {
        var orderedRoots = (0, xmlParser_js_1.parseXmlPreserveOrder)(rawXml);
        var findChildInOrdered = function (items) {
            var names = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                names[_i - 1] = arguments[_i];
            }
            for (var _a = 0, items_1 = items; _a < items_1.length; _a++) {
                var item = items_1[_a];
                for (var _b = 0, names_1 = names; _b < names_1.length; _b++) {
                    var name_2 = names_1[_b];
                    if (name_2 in item && Array.isArray(item[name_2])) {
                        return item[name_2];
                    }
                }
            }
            return undefined;
        };
        var docChildren = (_a = findChildInOrdered(orderedRoots, 'w:document', 'document')) !== null && _a !== void 0 ? _a : orderedRoots;
        var bodyChildren = findChildInOrdered(docChildren, 'w:body', 'body');
        if (bodyChildren && bodyChildren.length > 0) {
            var pIndex = 0;
            var tblIndex = 0;
            var sdtIndex = 0;
            for (var _i = 0, bodyChildren_1 = bodyChildren; _i < bodyChildren_1.length; _i++) {
                var child = bodyChildren_1[_i];
                var tagName = Object.keys(child).find(function (k) { return k !== ':@'; });
                if (!tagName)
                    continue;
                if (tagName === 'w:p' || tagName === 'p') {
                    if (pIndex < pNodes.length) {
                        var para = extractParagraph(pNodes[pIndex++], options, imageMap, styleMap);
                        orderedParagraphs.push(para);
                        bodyItems.push({ type: 'paragraph', paragraph: para });
                    }
                }
                else if (tagName === 'w:tbl' || tagName === 'tbl') {
                    if (tblIndex < tblNodes.length) {
                        var tableWrapper = { 'w:tbl': [tblNodes[tblIndex++]] };
                        var extracted = (0, tableExtractor_js_1.extractTables)(tableWrapper, options, imageMap, styleMap);
                        if (extracted.length > 0) {
                            orderedTables.push(extracted[0]);
                            bodyItems.push({ type: 'table', table: extracted[0] });
                        }
                    }
                }
                else if (tagName === 'w:sdt' || tagName === 'sdt') {
                    if (sdtIndex < sdtNodes.length) {
                        var sdtNode = sdtNodes[sdtIndex++];
                        var sdtContent = getProperty(sdtNode, 'w:sdtContent', 'sdtContent');
                        if (isRecord(sdtContent)) {
                            var sdtPNodes = toArray(getProperty(sdtContent, 'w:p', 'p'));
                            for (var _b = 0, sdtPNodes_1 = sdtPNodes; _b < sdtPNodes_1.length; _b++) {
                                var sdtP = sdtPNodes_1[_b];
                                var para = extractParagraph(sdtP, options, imageMap, styleMap);
                                orderedParagraphs.push(para);
                                bodyItems.push({ type: 'paragraph', paragraph: para });
                            }
                            var sdtTblNodes = toArray(getProperty(sdtContent, 'w:tbl', 'tbl'));
                            for (var _c = 0, sdtTblNodes_1 = sdtTblNodes; _c < sdtTblNodes_1.length; _c++) {
                                var sdtTbl = sdtTblNodes_1[_c];
                                var tableWrapper = { 'w:tbl': [sdtTbl] };
                                var extracted = (0, tableExtractor_js_1.extractTables)(tableWrapper, options, imageMap, styleMap);
                                if (extracted.length > 0) {
                                    orderedTables.push(extracted[0]);
                                    bodyItems.push({ type: 'table', table: extracted[0] });
                                }
                            }
                        }
                    }
                }
            }
            // Append any remaining paragraphs or tables if not visited in bodyChildren
            while (pIndex < pNodes.length) {
                var para = extractParagraph(pNodes[pIndex++], options, imageMap, styleMap);
                orderedParagraphs.push(para);
                bodyItems.push({ type: 'paragraph', paragraph: para });
            }
            while (tblIndex < tblNodes.length) {
                var tableWrapper = { 'w:tbl': [tblNodes[tblIndex++]] };
                var extracted = (0, tableExtractor_js_1.extractTables)(tableWrapper, options, imageMap, styleMap);
                if (extracted.length > 0) {
                    orderedTables.push(extracted[0]);
                    bodyItems.push({ type: 'table', table: extracted[0] });
                }
            }
            preserveOrderSuccess = true;
        }
    }
    catch (_d) {
        // Graceful fallback
    }
    if (!preserveOrderSuccess) {
        orderedParagraphs = pNodes.map(function (pNode) {
            return extractParagraph(pNode, options, imageMap, styleMap);
        });
        orderedTables = isRecord(bodyObj) ? (0, tableExtractor_js_1.extractTables)(bodyObj, options, imageMap, styleMap) : [];
        bodyItems = __spreadArray(__spreadArray([], orderedParagraphs.map(function (p) { return ({ type: 'paragraph', paragraph: p }); }), true), orderedTables.map(function (t) { return ({ type: 'table', table: t }); }), true);
    }
    var text = orderedParagraphs.map(function (p) { return p.text; }).join('\n');
    return __assign(__assign({ paragraphs: orderedParagraphs, text: text, tables: orderedTables }, (sections.length > 0 ? { sections: sections } : {})), { bodyItems: bodyItems });
}
/**
 * Resolves list item properties (abstractNumId and listType) across all paragraphs
 * in a DocxDocument using numbering definition maps.
 *
 * This function creates a new DocxDocument without mutating the input document.
 * For each paragraph with a listItem placeholder, it looks up abstractNumId from numIdMap,
 * and listType from abstractNumMap using key "${abstractNumId}:${level}".
 * If no matching list type is found, it defaults to 'bullet'.
 *
 * @param doc - The parsed DOCX document structure.
 * @param numIdMap - Mapping from numId to abstractNumId.
 * @param abstractNumMap - Mapping from "${abstractNumId}:${level}" to 'bullet' | 'ordered'.
 * @param abstractNumFmtMap - Optional mapping from "${abstractNumId}:${level}" to numFmt string.
 * @returns A new DocxDocument with resolved list items.
 */
function resolveListItems(doc, numIdMap, abstractNumMap, abstractNumFmtMap) {
    var _a;
    var resolvePara = function (para) {
        if (!para.listItem) {
            return __assign({}, para);
        }
        var _a = para.listItem, numId = _a.numId, level = _a.level;
        var abstractNumId = numIdMap.get(numId);
        if (abstractNumId === undefined) {
            if (abstractNumMap.has("".concat(numId, ":").concat(level)) || abstractNumMap.has("".concat(numId, ":0"))) {
                abstractNumId = numId;
            }
            else {
                abstractNumId = 0;
            }
        }
        var listType = abstractNumMap.get("".concat(abstractNumId, ":").concat(level));
        var numFmt = abstractNumFmtMap === null || abstractNumFmtMap === void 0 ? void 0 : abstractNumFmtMap.get("".concat(abstractNumId, ":").concat(level));
        if (listType === undefined) {
            for (var l = level - 1; l >= 0; l--) {
                var fallbackType = abstractNumMap.get("".concat(abstractNumId, ":").concat(l));
                if (fallbackType !== undefined) {
                    listType = fallbackType;
                    numFmt = abstractNumFmtMap === null || abstractNumFmtMap === void 0 ? void 0 : abstractNumFmtMap.get("".concat(abstractNumId, ":").concat(l));
                    break;
                }
            }
        }
        if (listType === undefined) {
            for (var _i = 0, _b = abstractNumMap.entries(); _i < _b.length; _i++) {
                var _c = _b[_i], key = _c[0], type = _c[1];
                if (key.startsWith("".concat(abstractNumId, ":"))) {
                    listType = type;
                    numFmt = abstractNumFmtMap === null || abstractNumFmtMap === void 0 ? void 0 : abstractNumFmtMap.get(key);
                    break;
                }
            }
        }
        if (listType === undefined) {
            listType = 'bullet';
        }
        var listItem = __assign({ abstractNumId: abstractNumId, numId: numId, level: level, listType: listType }, (numFmt !== undefined ? { numFmt: numFmt } : {}));
        return __assign(__assign({}, para), { listItem: listItem });
    };
    var paragraphs = doc.paragraphs.map(resolvePara);
    var bodyItems = (_a = doc.bodyItems) === null || _a === void 0 ? void 0 : _a.map(function (item) {
        if (item.type === 'paragraph') {
            return {
                type: 'paragraph',
                paragraph: resolvePara(item.paragraph),
            };
        }
        return item;
    });
    return __assign(__assign(__assign({}, doc), { paragraphs: paragraphs }), (bodyItems ? { bodyItems: bodyItems } : {}));
}
