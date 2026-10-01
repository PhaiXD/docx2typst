"use strict";
/**
 * @file imageExtractor.ts
 * Module for extracting images and image metadata from OOXML DrawingML structures.
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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
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
exports.findImagesInRun = findImagesInRun;
exports.extractImageMetadata = extractImageMetadata;
exports.saveImages = saveImages;
var node_path_1 = require("node:path");
var promises_1 = require("node:fs/promises");
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
 * Extracts relationship IDs from a drawing container node (`wp:inline` or `wp:anchor`).
 *
 * Navigates: container -> `a:graphic` -> `a:graphicData` -> `pic:pic` -> `pic:blipFill` -> `a:blip @r:embed`.
 *
 * @param container - The inline or anchor drawing container node.
 * @returns Array of relationship ID strings found in the container.
 */
function extractRelIdsFromContainer(container) {
    var relIds = [];
    var graphicNodes = toArray(getProperty(container, 'a:graphic', 'graphic'));
    for (var _i = 0, graphicNodes_1 = graphicNodes; _i < graphicNodes_1.length; _i++) {
        var graphic = graphicNodes_1[_i];
        if (!isRecord(graphic)) {
            continue;
        }
        var graphicDataNodes = toArray(getProperty(graphic, 'a:graphicData', 'graphicData'));
        for (var _a = 0, graphicDataNodes_1 = graphicDataNodes; _a < graphicDataNodes_1.length; _a++) {
            var graphicData = graphicDataNodes_1[_a];
            if (!isRecord(graphicData)) {
                continue;
            }
            var picNodes = toArray(getProperty(graphicData, 'pic:pic', 'pic'));
            for (var _b = 0, picNodes_1 = picNodes; _b < picNodes_1.length; _b++) {
                var pic = picNodes_1[_b];
                if (!isRecord(pic)) {
                    continue;
                }
                var blipFillNodes = toArray(getProperty(pic, 'pic:blipFill', 'blipFill'));
                for (var _c = 0, blipFillNodes_1 = blipFillNodes; _c < blipFillNodes_1.length; _c++) {
                    var blipFill = blipFillNodes_1[_c];
                    if (!isRecord(blipFill)) {
                        continue;
                    }
                    var blipNodes = toArray(getProperty(blipFill, 'a:blip', 'blip'));
                    for (var _d = 0, blipNodes_1 = blipNodes; _d < blipNodes_1.length; _d++) {
                        var blip = blipNodes_1[_d];
                        if (!isRecord(blip)) {
                            continue;
                        }
                        var rEmbed = getProperty(blip, '@_r:embed', '@_embed', '@r:embed', '@embed', 'r:embed', 'embed');
                        if (rEmbed !== undefined && rEmbed !== null) {
                            var strVal = String(rEmbed).trim();
                            if (strVal.length > 0) {
                                relIds.push(strVal);
                            }
                        }
                    }
                }
            }
        }
    }
    return relIds;
}
/**
 * Searches a run node (`w:r`) for embedded image drawing relationships.
 *
 * Looks for `w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill > a:blip @r:embed`.
 * Also checks `w:drawing > wp:anchor` as an alternative to `wp:inline`.
 * Navigates properties using both namespace and non-namespace keys.
 *
 * @param rNode - Parsed run node representation.
 * @returns Array of relationship IDs found (may be empty).
 */
function findImagesInRun(rNode) {
    if (!isRecord(rNode)) {
        return [];
    }
    var drawingNodes = toArray(getProperty(rNode, 'w:drawing', 'drawing'));
    var relIds = [];
    for (var _i = 0, drawingNodes_1 = drawingNodes; _i < drawingNodes_1.length; _i++) {
        var drawing = drawingNodes_1[_i];
        if (!isRecord(drawing)) {
            continue;
        }
        var inlines = toArray(getProperty(drawing, 'wp:inline', 'inline'));
        var anchors = toArray(getProperty(drawing, 'wp:anchor', 'anchor'));
        var containers = __spreadArray(__spreadArray([], inlines, true), anchors, true);
        for (var _a = 0, containers_1 = containers; _a < containers_1.length; _a++) {
            var container = containers_1[_a];
            if (!isRecord(container)) {
                continue;
            }
            relIds.push.apply(relIds, extractRelIdsFromContainer(container));
        }
    }
    return relIds;
}
/**
 * Extracts image metadata for images found within a run node.
 *
 * Calls `findImagesInRun` to discover relationship IDs, resolves each against `imageMap`,
 * and extracts dimension (`wp:extent @cx, @cy`) and alt text (`wp:docPr @descr` or `@title`)
 * metadata. Returns deep copies without mutating the `imageMap` entries.
 *
 * @param rNode - Parsed run node representation.
 * @param imageMap - Mapping from relationship ID to base DocxImage metadata.
 * @returns Array of DocxImage objects with resolved dimensions and alt text.
 */
function extractImageMetadata(rNode, imageMap) {
    if (!isRecord(rNode) || !imageMap || imageMap.size === 0) {
        return [];
    }
    var foundRelIds = findImagesInRun(rNode);
    if (foundRelIds.length === 0) {
        return [];
    }
    var drawingNodes = toArray(getProperty(rNode, 'w:drawing', 'drawing'));
    var result = [];
    for (var _i = 0, drawingNodes_2 = drawingNodes; _i < drawingNodes_2.length; _i++) {
        var drawing = drawingNodes_2[_i];
        if (!isRecord(drawing)) {
            continue;
        }
        var inlines = toArray(getProperty(drawing, 'wp:inline', 'inline'));
        var anchors = toArray(getProperty(drawing, 'wp:anchor', 'anchor'));
        var containers = __spreadArray(__spreadArray([], inlines, true), anchors, true);
        for (var _a = 0, containers_2 = containers; _a < containers_2.length; _a++) {
            var container = containers_2[_a];
            if (!isRecord(container)) {
                continue;
            }
            var relIds = extractRelIdsFromContainer(container);
            if (relIds.length === 0) {
                continue;
            }
            // Extract width and height from wp:extent @cx and @cy
            var extentNode = getProperty(container, 'wp:extent', 'extent');
            var widthEmu = void 0;
            var heightEmu = void 0;
            if (isRecord(extentNode)) {
                var cx = getProperty(extentNode, '@_cx', '@cx', 'cx');
                if (cx !== undefined && cx !== null) {
                    var parsedCx = parseInt(String(cx), 10);
                    if (!isNaN(parsedCx)) {
                        widthEmu = parsedCx;
                    }
                }
                var cy = getProperty(extentNode, '@_cy', '@cy', 'cy');
                if (cy !== undefined && cy !== null) {
                    var parsedCy = parseInt(String(cy), 10);
                    if (!isNaN(parsedCy)) {
                        heightEmu = parsedCy;
                    }
                }
            }
            // Extract altText from wp:docPr @descr or @title
            var docPrNode = getProperty(container, 'wp:docPr', 'docPr');
            var altText = void 0;
            if (isRecord(docPrNode)) {
                var descr = getProperty(docPrNode, '@_descr', '@descr', 'descr');
                var title = getProperty(docPrNode, '@_title', '@title', 'title');
                if (descr !== undefined && descr !== null && String(descr).trim().length > 0) {
                    altText = String(descr);
                }
                else if (title !== undefined && title !== null && String(title).trim().length > 0) {
                    altText = String(title);
                }
            }
            for (var _b = 0, relIds_1 = relIds; _b < relIds_1.length; _b++) {
                var relId = relIds_1[_b];
                var baseImage = imageMap.get(relId);
                if (!baseImage) {
                    continue;
                }
                // Deep copy the base DocxImage to avoid mutating imageMap entries
                var imageCopy = __assign(__assign(__assign(__assign({}, baseImage), (widthEmu !== undefined ? { widthEmu: widthEmu } : {})), (heightEmu !== undefined ? { heightEmu: heightEmu } : {})), (altText !== undefined ? { altText: altText } : {}));
                result.push(imageCopy);
            }
        }
    }
    return result;
}
/**
 * Extracts image binary data from a DOCX zip archive and writes files to the filesystem.
 *
 * @param images - Array of DocxImage items to save.
 * @param zipInstance - The JSZip archive instance, typed as unknown and cast internally.
 * @param outputDir - Destination directory path for extracted images.
 * @returns A promise resolving to a Map of `zipPath` to saved file path.
 */
function saveImages(images, zipInstance, outputDir) {
    return __awaiter(this, void 0, void 0, function () {
        var savedMap, zip, _i, images_1, image, fileEntry, data, filename, cleanOutputDir, savedPath;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    savedMap = new Map();
                    if (!images || images.length === 0 || !zipInstance) {
                        return [2 /*return*/, savedMap];
                    }
                    zip = zipInstance;
                    if (typeof zip.file !== 'function') {
                        return [2 /*return*/, savedMap];
                    }
                    if (!(outputDir.trim().length > 0)) return [3 /*break*/, 2];
                    return [4 /*yield*/, (0, promises_1.mkdir)(outputDir, { recursive: true })];
                case 1:
                    _a.sent();
                    _a.label = 2;
                case 2:
                    _i = 0, images_1 = images;
                    _a.label = 3;
                case 3:
                    if (!(_i < images_1.length)) return [3 /*break*/, 7];
                    image = images_1[_i];
                    fileEntry = zip.file(image.zipPath);
                    if (!fileEntry) {
                        return [3 /*break*/, 6];
                    }
                    return [4 /*yield*/, fileEntry.async('uint8array')];
                case 4:
                    data = _a.sent();
                    filename = (0, node_path_1.basename)(image.zipPath.replace(/\\/g, '/'));
                    cleanOutputDir = outputDir.replace(/[/\\]+$/, '');
                    savedPath = cleanOutputDir.length > 0 ? "".concat(cleanOutputDir, "/").concat(filename) : filename;
                    return [4 /*yield*/, (0, promises_1.writeFile)(savedPath, data)];
                case 5:
                    _a.sent();
                    savedMap.set(image.zipPath, savedPath);
                    _a.label = 6;
                case 6:
                    _i++;
                    return [3 /*break*/, 3];
                case 7: return [2 /*return*/, savedMap];
            }
        });
    });
}
