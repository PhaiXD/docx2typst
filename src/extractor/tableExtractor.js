"use strict";
/**
 * @file tableExtractor.ts
 * Module for extracting structured tables (w:tbl) from OOXML WordprocessingML content.
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.extractTables = extractTables;
var textExtractor_js_1 = require("./textExtractor.js");
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
 * Checks whether a TableCell contains no meaningful text or images.
 *
 * @param cell - Target table cell.
 * @returns True if the cell is completely empty.
 */
function isCellEmpty(cell) {
    if (!cell.paragraphs || cell.paragraphs.length === 0) {
        return true;
    }
    return cell.paragraphs.every(function (p) {
        var hasText = p.text && p.text.trim().length > 0;
        var hasImages = Array.isArray(p.images) && p.images.length > 0;
        return !hasText && !hasImages;
    });
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
 * Extracts all tables (`w:tbl`) from a parsed WordprocessingML body node.
 *
 * Navigates the OOXML hierarchy: `w:tbl` -> `w:tr` -> `w:tc` -> `w:p`.
 * Extracts column count, row structure, cell spans, vertical merge indicators,
 * and detects header rows based on paragraph styles.
 *
 * @param bodyNode - The parsed `w:body` node from the document XML tree.
 * @returns An array of structured DocxTable objects.
 */
function extractTables(bodyNode, options, imageMap, styleMap) {
    var _a, _b;
    if (!isRecord(bodyNode)) {
        return [];
    }
    var tblNodes = toArray(getProperty(bodyNode, 'w:tbl', 'tbl'));
    var tables = [];
    var _loop_1 = function (tblNode) {
        if (!isRecord(tblNode)) {
            return "continue";
        }
        // Extract table style from w:tblPr > w:tblStyle @w:val
        var style = void 0;
        var tblPr = getProperty(tblNode, 'w:tblPr', 'tblPr');
        if (isRecord(tblPr)) {
            var tblStyle = getProperty(tblPr, 'w:tblStyle', 'tblStyle');
            if (isRecord(tblStyle)) {
                var val = getProperty(tblStyle, '@_w:val', '@_val', '@w:val', 'val');
                if (val !== undefined && val !== null) {
                    style = String(val);
                }
            }
            else if (typeof tblStyle === 'string') {
                style = tblStyle;
            }
        }
        // Extract table borders from w:tblPr > w:tblBorders
        var borders = void 0;
        if (isRecord(tblPr)) {
            var tblBordersNode_1 = getProperty(tblPr, 'w:tblBorders', 'tblBorders');
            if (isRecord(tblBordersNode_1)) {
                var checkBorder = function (nodeKey, aliasKey) {
                    var borderNode = getProperty(tblBordersNode_1, nodeKey, aliasKey);
                    if (isRecord(borderNode)) {
                        var val = getProperty(borderNode, '@_w:val', '@_val', '@w:val', 'val');
                        if (val !== undefined && val !== null) {
                            var valStr = String(val).toLowerCase().trim();
                            return valStr !== 'none' && valStr !== 'nil';
                        }
                    }
                    return false;
                };
                borders = {
                    top: checkBorder('w:top', 'top'),
                    bottom: checkBorder('w:bottom', 'bottom'),
                    left: checkBorder('w:left', 'left'),
                    right: checkBorder('w:right', 'right'),
                    insideH: checkBorder('w:insideH', 'insideH'),
                    insideV: checkBorder('w:insideV', 'insideV'),
                };
            }
        }
        // Rows in table
        var trNodes = toArray(getProperty(tblNode, 'w:tr', 'tr'));
        var hasAnyCellBorders = false;
        var hasCellExplicitlyRemovedVertical = false;
        var hasAnyCellVerticalBorder = false;
        // Determine columnCount: from w:tblGrid > w:gridCol OR first row gridSpan sum
        var columnCount = 0;
        var tblGrid = getProperty(tblNode, 'w:tblGrid', 'tblGrid');
        if (isRecord(tblGrid)) {
            var gridCols = toArray(getProperty(tblGrid, 'w:gridCol', 'gridCol'));
            if (gridCols.length > 0) {
                columnCount = gridCols.length;
            }
        }
        if (columnCount === 0 && trNodes.length > 0) {
            for (var _c = 0, trNodes_1 = trNodes; _c < trNodes_1.length; _c++) {
                var r = trNodes_1[_c];
                if (isRecord(r)) {
                    var tcList = toArray(getProperty(r, 'w:tc', 'tc'));
                    var rowCols = 0;
                    for (var _d = 0, tcList_1 = tcList; _d < tcList_1.length; _d++) {
                        var tc = tcList_1[_d];
                        var span = 1;
                        if (isRecord(tc)) {
                            var tcPr = getProperty(tc, 'w:tcPr', 'tcPr');
                            if (isRecord(tcPr)) {
                                var gs = getProperty(tcPr, 'w:gridSpan', 'gridSpan');
                                if (isRecord(gs)) {
                                    var val = getProperty(gs, '@_w:val', '@_val', '@w:val', 'val');
                                    var parsed = parseInt(String(val), 10);
                                    if (!isNaN(parsed))
                                        span = parsed;
                                }
                                else if (typeof gs === 'number') {
                                    span = Math.trunc(gs);
                                }
                                else if (typeof gs === 'string') {
                                    var parsed = parseInt(gs, 10);
                                    if (!isNaN(parsed))
                                        span = parsed;
                                }
                            }
                        }
                        rowCols += span;
                    }
                    if (rowCols > 0) {
                        columnCount = rowCols;
                        break;
                    }
                }
            }
        }
        // Extract column widths from w:tblGrid > w:gridCol
        var columnWidths = void 0;
        if (isRecord(tblGrid)) {
            var gridCols = toArray(getProperty(tblGrid, 'w:gridCol', 'gridCol'));
            var widths = [];
            for (var _e = 0, gridCols_1 = gridCols; _e < gridCols_1.length; _e++) {
                var gc_1 = gridCols_1[_e];
                if (isRecord(gc_1)) {
                    var w = getProperty(gc_1, '@_w:w', '@_w', '@w:w', 'w');
                    if (w !== undefined && w !== null) {
                        var parsed = parseInt(String(w), 10);
                        widths.push(!isNaN(parsed) ? parsed : 0);
                    }
                    else {
                        widths.push(0);
                    }
                }
            }
            if (widths.length > 0) {
                columnWidths = widths;
            }
        }
        var rows = [];
        for (var _f = 0, trNodes_2 = trNodes; _f < trNodes_2.length; _f++) {
            var trNode = trNodes_2[_f];
            if (!isRecord(trNode)) {
                continue;
            }
            var tcNodes = toArray(getProperty(trNode, 'w:tc', 'tc'));
            var cells = [];
            var consumed = 0;
            for (var _g = 0, tcNodes_1 = tcNodes; _g < tcNodes_1.length; _g++) {
                var tcNode = tcNodes_1[_g];
                if (!isRecord(tcNode)) {
                    continue;
                }
                var tcPr = getProperty(tcNode, 'w:tcPr', 'tcPr');
                var columnSpan = void 0;
                var isVerticalMerge = void 0;
                var vMerge = void 0;
                var cellBorders = void 0;
                if (isRecord(tcPr)) {
                    // Extract columnSpan from w:tcPr > w:gridSpan @w:val (default 1 if missing)
                    var gridSpanNode = getProperty(tcPr, 'w:gridSpan', 'gridSpan');
                    if (gridSpanNode !== undefined) {
                        if (isRecord(gridSpanNode)) {
                            var val = getProperty(gridSpanNode, '@_w:val', '@_val', '@w:val', 'val');
                            if (val !== undefined && val !== null) {
                                var parsed = parseInt(String(val), 10);
                                columnSpan = !isNaN(parsed) ? parsed : 1;
                            }
                            else {
                                columnSpan = 1;
                            }
                        }
                        else if (typeof gridSpanNode === 'number') {
                            columnSpan = Math.trunc(gridSpanNode);
                        }
                        else if (typeof gridSpanNode === 'string') {
                            var parsed = parseInt(gridSpanNode, 10);
                            columnSpan = !isNaN(parsed) ? parsed : 1;
                        }
                        else {
                            columnSpan = 1;
                        }
                    }
                    // Extract isVerticalMerge: if w:tcPr > w:vMerge element exists
                    var vMergeNode = getProperty(tcPr, 'w:vMerge', 'vMerge');
                    if (vMergeNode !== undefined) {
                        isVerticalMerge = true;
                        var val = void 0;
                        if (isRecord(vMergeNode)) {
                            var rawVal = getProperty(vMergeNode, '@_w:val', '@_val', '@w:val', 'val');
                            if (rawVal !== undefined && rawVal !== null) {
                                val = String(rawVal).toLowerCase().trim();
                            }
                        }
                        else if (typeof vMergeNode === 'string') {
                            val = vMergeNode.toLowerCase().trim();
                        }
                        if (val === 'restart') {
                            vMerge = 'restart';
                        }
                        else {
                            vMerge = 'continue';
                        }
                    }
                    // Check cell borders
                    var tcBorders = getProperty(tcPr, 'w:tcBorders', 'tcBorders');
                    if (isRecord(tcBorders)) {
                        hasAnyCellBorders = true;
                        cellBorders = {};
                        for (var _h = 0, _j = ['top', 'bottom', 'left', 'right']; _h < _j.length; _h++) {
                            var side = _j[_h];
                            var bNode = getProperty(tcBorders, "w:".concat(side), side);
                            if (isRecord(bNode)) {
                                var val = getProperty(bNode, '@_w:val', '@_val', '@w:val', 'val');
                                if (val !== undefined && val !== null) {
                                    var s = String(val).toLowerCase().trim();
                                    if (s === 'none' || s === 'nil') {
                                        cellBorders[side] = false;
                                        if (side === 'left' || side === 'right')
                                            hasCellExplicitlyRemovedVertical = true;
                                    }
                                    else {
                                        cellBorders[side] = true;
                                        if (side === 'left' || side === 'right')
                                            hasAnyCellVerticalBorder = true;
                                    }
                                }
                            }
                        }
                        if (Object.keys(cellBorders).length === 0) {
                            cellBorders = undefined;
                        }
                    }
                }
                // Extract cell paragraphs
                var pNodes = toArray(getProperty(tcNode, 'w:p', 'p'));
                var paragraphs = pNodes.map(function (pNode) {
                    return (0, textExtractor_js_1.extractParagraph)(pNode, options, imageMap, styleMap);
                });
                // Inherit table style spacing for cell paragraphs if not explicitly defined
                var tableStyleInfo = style ? styleMap === null || styleMap === void 0 ? void 0 : styleMap.get(style) : undefined;
                if (tableStyleInfo === null || tableStyleInfo === void 0 ? void 0 : tableStyleInfo.spacing) {
                    for (var _k = 0, paragraphs_1 = paragraphs; _k < paragraphs_1.length; _k++) {
                        var p = paragraphs_1[_k];
                        if (!p.spacing) {
                            p.spacing = __assign({}, tableStyleInfo.spacing);
                        }
                    }
                }
                // Extract vertical alignment from w:tcPr > w:vAlign @w:val
                var align = void 0;
                if (isRecord(tcPr)) {
                    var vAlignNode = getProperty(tcPr, 'w:vAlign', 'vAlign');
                    if (vAlignNode !== undefined) {
                        var vAlignVal = void 0;
                        if (isRecord(vAlignNode)) {
                            var val = getProperty(vAlignNode, '@_w:val', '@_val', '@w:val', 'val');
                            if (val !== undefined && val !== null) {
                                vAlignVal = String(val).toLowerCase().trim();
                            }
                        }
                        else if (typeof vAlignNode === 'string') {
                            vAlignVal = vAlignNode.toLowerCase().trim();
                        }
                        if (vAlignVal === 'center') {
                            align = 'horizon';
                        }
                        else if (vAlignVal === 'top') {
                            align = 'top';
                        }
                        else if (vAlignVal === 'bottom') {
                            align = 'bottom';
                        }
                    }
                }
                var cell = __assign(__assign(__assign(__assign(__assign({ paragraphs: paragraphs }, (columnSpan !== undefined ? { columnSpan: columnSpan } : {})), (isVerticalMerge ? { isVerticalMerge: true } : {})), (vMerge !== undefined ? { vMerge: vMerge } : {})), (align !== undefined ? { align: align } : {})), (cellBorders !== undefined ? { borders: cellBorders } : {}));
                // If logical cells already satisfy columnCount, skip extra phantom cells
                if (columnCount > 0 && consumed >= columnCount && isCellEmpty(cell)) {
                    break;
                }
                cells.push(cell);
                consumed += columnSpan !== null && columnSpan !== void 0 ? columnSpan : 1;
            }
            // Drop trailing phantom cells (isVerticalMerge=true and isEmpty) if exceeding columnCount
            if (columnCount > 0) {
                while (cells.length > 0) {
                    var totalSlots = cells.reduce(function (sum, c) { var _a; return sum + ((_a = c.columnSpan) !== null && _a !== void 0 ? _a : 1); }, 0);
                    if (totalSlots <= columnCount) {
                        break;
                    }
                    var lastCell = cells[cells.length - 1];
                    if (lastCell.isVerticalMerge && isCellEmpty(lastCell)) {
                        cells.pop();
                    }
                    else {
                        break;
                    }
                }
            }
            // Detect isHeader for the row: if any cell's first paragraph has a style containing 'heading' or 'header'
            var isHeader = false;
            for (var _l = 0, cells_1 = cells; _l < cells_1.length; _l++) {
                var cell = cells_1[_l];
                var firstPara = cell.paragraphs[0];
                if ((firstPara === null || firstPara === void 0 ? void 0 : firstPara.style) && /heading|header/i.test(firstPara.style)) {
                    isHeader = true;
                    break;
                }
            }
            var row = __assign({ cells: cells }, (isHeader ? { isHeader: true } : {}));
            rows.push(row);
        }
        // Calculate rowSpan for vMerge='restart' cells
        for (var r = 0; r < rows.length; r++) {
            var colIdx = 0;
            for (var _m = 0, _o = rows[r].cells; _m < _o.length; _m++) {
                var cell = _o[_m];
                var span = (_a = cell.columnSpan) !== null && _a !== void 0 ? _a : 1;
                if (cell.vMerge === 'restart') {
                    var rowSpan = 1;
                    for (var nextR = r + 1; nextR < rows.length; nextR++) {
                        var nextCol = 0;
                        var matchingCell = void 0;
                        for (var _p = 0, _q = rows[nextR].cells; _p < _q.length; _p++) {
                            var c = _q[_p];
                            if (nextCol === colIdx) {
                                matchingCell = c;
                                break;
                            }
                            nextCol += (_b = c.columnSpan) !== null && _b !== void 0 ? _b : 1;
                        }
                        if (matchingCell && matchingCell.vMerge === 'continue') {
                            rowSpan++;
                        }
                        else {
                            break;
                        }
                    }
                    if (rowSpan > 1) {
                        cell.rowSpan = rowSpan;
                    }
                }
                colIdx += span;
            }
        }
        var hasBorderTop = false;
        var hasBorderBottom = false;
        if (rows.length > 0) {
            var firstRow = rows[0];
            if (firstRow.cells.some(function (cell) { return cell.paragraphs.some(function (p) { return p.hasBorderTop; }); })) {
                hasBorderTop = true;
            }
        }
        if (hasBorderTop && trNodes.length > 0) {
            var lastTrNode = trNodes[trNodes.length - 1];
            if (isRecord(lastTrNode)) {
                var lastTcNodes = toArray(getProperty(lastTrNode, 'w:tc', 'tc'));
                for (var _r = 0, lastTcNodes_1 = lastTcNodes; _r < lastTcNodes_1.length; _r++) {
                    var tc = lastTcNodes_1[_r];
                    if (isRecord(tc)) {
                        var tcPr = getProperty(tc, 'w:tcPr', 'tcPr');
                        if (isRecord(tcPr)) {
                            var tcBorders = getProperty(tcPr, 'w:tcBorders', 'tcBorders');
                            if (isRecord(tcBorders)) {
                                var bottomBdr = getProperty(tcBorders, 'w:bottom', 'bottom');
                                if (isRecord(bottomBdr)) {
                                    var val = getProperty(bottomBdr, '@_w:val', '@_val', '@w:val', 'val');
                                    if (typeof val === 'string' && val.toLowerCase() !== 'none' && val.toLowerCase() !== 'nil') {
                                        hasBorderBottom = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (!hasBorderBottom && rows.length > 0) {
                var lastRow = rows[rows.length - 1];
                if (lastRow.cells.some(function (cell) { return cell.paragraphs.some(function (p) { return p.hasBorderBottom; }); })) {
                    hasBorderBottom = true;
                }
            }
        }
        if (!borders && hasAnyCellBorders) {
            borders = {
                top: hasBorderTop,
                bottom: hasBorderBottom,
                left: hasAnyCellVerticalBorder,
                right: hasAnyCellVerticalBorder,
                insideH: false,
                insideV: hasAnyCellVerticalBorder,
            };
        }
        var table = __assign(__assign(__assign(__assign(__assign({ rows: rows, columnCount: columnCount }, (columnWidths && columnWidths.length === columnCount ? { columnWidths: columnWidths } : {})), (style !== undefined ? { style: style } : {})), (borders !== undefined ? { borders: borders } : {})), (hasBorderTop ? { hasBorderTop: true } : {})), (hasBorderBottom ? { hasBorderBottom: true } : {}));
        tables.push(table);
    };
    for (var _i = 0, tblNodes_1 = tblNodes; _i < tblNodes_1.length; _i++) {
        var tblNode = tblNodes_1[_i];
        _loop_1(tblNode);
    }
    return tables;
}
