"use strict";
/**
 * @file reader.ts
 * Module for loading .docx archives from disk and extracting raw XML parts.
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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.readDocxFile = readDocxFile;
var promises_1 = require("node:fs/promises");
var jszip_1 = __importDefault(require("jszip"));
var errors_js_1 = require("./errors.js");
var xmlParser_js_1 = require("./utils/xmlParser.js");
/**
 * Path to the main document body XML inside the .docx ZIP package.
 */
var DOCUMENT_XML_PATH = 'word/document.xml';
/**
 * Path to the document relationships XML inside the .docx ZIP package.
 */
var DOCUMENT_RELS_XML_PATH = 'word/_rels/document.xml.rels';
/**
 * Path to the numbering definitions XML inside the .docx ZIP package.
 */
var NUMBERING_XML_PATH = 'word/numbering.xml';
/**
 * Path to the styles definitions XML inside the .docx ZIP package.
 */
var STYLES_XML_PATH = 'word/styles.xml';
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
 * Detects the MIME type of an image file based on its file extension.
 *
 * @param filePath - The path or filename of the image.
 * @returns The corresponding MIME type string, defaulting to 'image/png'.
 */
function detectMimeType(filePath) {
    var _a;
    var ext = (_a = filePath.toLowerCase().split('.').pop()) !== null && _a !== void 0 ? _a : '';
    switch (ext) {
        case 'png':
            return 'image/png';
        case 'jpg':
        case 'jpeg':
            return 'image/jpeg';
        case 'gif':
            return 'image/gif';
        case 'webp':
            return 'image/webp';
        case 'svg':
            return 'image/svg+xml';
        case 'emf':
            return 'image/emf';
        case 'wmf':
            return 'image/wmf';
        default:
            return 'image/png';
    }
}
/**
 * Parses OOXML relationships XML (`word/_rels/document.xml.rels`) to extract target mappings.
 *
 * @param relsXml - Raw XML string from the relationships part.
 * @returns Map of relationship ID to target path and relationship type.
 */
function parseRelationshipsForImages(relsXml) {
    var map = new Map();
    if (typeof relsXml !== 'string' || relsXml.trim().length === 0) {
        return map;
    }
    var parsed;
    try {
        parsed = (0, xmlParser_js_1.parseXml)(relsXml);
    }
    catch (_a) {
        return map;
    }
    var relsRoot = getProperty(parsed, 'Relationships', 'relationships', 'r:Relationships');
    var rootObj = isRecord(relsRoot) ? relsRoot : parsed;
    var relNodes = toArray(getProperty(rootObj, 'Relationship', 'relationship', 'r:Relationship'));
    for (var _i = 0, relNodes_1 = relNodes; _i < relNodes_1.length; _i++) {
        var relNode = relNodes_1[_i];
        if (!isRecord(relNode)) {
            continue;
        }
        var id = getProperty(relNode, '@_Id', '@_id', '@Id', '@id', 'Id', 'id');
        var type = getProperty(relNode, '@_Type', '@_type', '@Type', '@type', 'Type', 'type');
        var target = getProperty(relNode, '@_Target', '@_target', '@Target', '@target', 'Target', 'target');
        if (id !== undefined && type !== undefined && target !== undefined) {
            map.set(String(id), {
                target: String(target),
                type: String(type),
            });
        }
    }
    return map;
}
/**
 * Reads a .docx file from the local filesystem and extracts its primary XML streams.
 *
 * @param filePath - The absolute or relative file path to the target .docx file.
 * @returns A promise resolving to the extracted raw XML content.
 * @throws {DocxReadError} If reading the file fails, if the file is not a valid ZIP archive,
 *                         or if the required `word/document.xml` part is missing.
 */
function readDocxFile(filePath) {
    return __awaiter(this, void 0, void 0, function () {
        var fileBuffer, error_1, errorMsg, zip, error_2, errorMsg, documentXmlEntry, documentXml, error_3, errorMsg, relationshipsXml, relsEntry, _a, numberingXml, numberingEntry, _b, stylesXml, stylesEntry, _c, imageMap, rels, _i, _d, _e, relationshipId, _f, target, type, zipPath, mimeType;
        return __generator(this, function (_g) {
            switch (_g.label) {
                case 0:
                    if (typeof filePath !== 'string' || filePath.trim().length === 0) {
                        throw new errors_js_1.DocxReadError('A non-empty file path string must be provided.');
                    }
                    _g.label = 1;
                case 1:
                    _g.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, (0, promises_1.readFile)(filePath)];
                case 2:
                    fileBuffer = _g.sent();
                    return [3 /*break*/, 4];
                case 3:
                    error_1 = _g.sent();
                    errorMsg = error_1 instanceof Error ? error_1.message : String(error_1);
                    throw new errors_js_1.DocxReadError("Failed to read file from disk at \"".concat(filePath, "\": ").concat(errorMsg), {
                        cause: error_1,
                    });
                case 4:
                    _g.trys.push([4, 6, , 7]);
                    return [4 /*yield*/, jszip_1.default.loadAsync(fileBuffer)];
                case 5:
                    zip = _g.sent();
                    return [3 /*break*/, 7];
                case 6:
                    error_2 = _g.sent();
                    errorMsg = error_2 instanceof Error ? error_2.message : String(error_2);
                    throw new errors_js_1.DocxReadError("File at \"".concat(filePath, "\" is not a valid ZIP/DOCX archive: ").concat(errorMsg), {
                        cause: error_2,
                    });
                case 7:
                    documentXmlEntry = zip.file(DOCUMENT_XML_PATH);
                    if (!documentXmlEntry) {
                        throw new errors_js_1.DocxReadError("Invalid .docx package at \"".concat(filePath, "\": missing required entry \"").concat(DOCUMENT_XML_PATH, "\"."));
                    }
                    _g.label = 8;
                case 8:
                    _g.trys.push([8, 10, , 11]);
                    return [4 /*yield*/, documentXmlEntry.async('string')];
                case 9:
                    documentXml = _g.sent();
                    return [3 /*break*/, 11];
                case 10:
                    error_3 = _g.sent();
                    errorMsg = error_3 instanceof Error ? error_3.message : String(error_3);
                    throw new errors_js_1.DocxReadError("Failed to extract \"".concat(DOCUMENT_XML_PATH, "\" from \"").concat(filePath, "\": ").concat(errorMsg), { cause: error_3 });
                case 11:
                    relsEntry = zip.file(DOCUMENT_RELS_XML_PATH);
                    if (!relsEntry) return [3 /*break*/, 15];
                    _g.label = 12;
                case 12:
                    _g.trys.push([12, 14, , 15]);
                    return [4 /*yield*/, relsEntry.async('string')];
                case 13:
                    relationshipsXml = _g.sent();
                    return [3 /*break*/, 15];
                case 14:
                    _a = _g.sent();
                    // Relationships are optional; proceed without relationships if read fails
                    relationshipsXml = undefined;
                    return [3 /*break*/, 15];
                case 15:
                    numberingEntry = zip.file(NUMBERING_XML_PATH);
                    if (!numberingEntry) return [3 /*break*/, 19];
                    _g.label = 16;
                case 16:
                    _g.trys.push([16, 18, , 19]);
                    return [4 /*yield*/, numberingEntry.async('string')];
                case 17:
                    numberingXml = _g.sent();
                    return [3 /*break*/, 19];
                case 18:
                    _b = _g.sent();
                    // Numbering is optional; proceed without numbering if read fails
                    numberingXml = undefined;
                    return [3 /*break*/, 19];
                case 19:
                    stylesEntry = zip.file(STYLES_XML_PATH);
                    if (!stylesEntry) return [3 /*break*/, 23];
                    _g.label = 20;
                case 20:
                    _g.trys.push([20, 22, , 23]);
                    return [4 /*yield*/, stylesEntry.async('string')];
                case 21:
                    stylesXml = _g.sent();
                    return [3 /*break*/, 23];
                case 22:
                    _c = _g.sent();
                    // Styles are optional; proceed without styles if read fails
                    stylesXml = undefined;
                    return [3 /*break*/, 23];
                case 23:
                    imageMap = new Map();
                    if (relationshipsXml) {
                        rels = parseRelationshipsForImages(relationshipsXml);
                        for (_i = 0, _d = rels.entries(); _i < _d.length; _i++) {
                            _e = _d[_i], relationshipId = _e[0], _f = _e[1], target = _f.target, type = _f.type;
                            if (type.endsWith('/image')) {
                                zipPath = target.startsWith('word/') ? target : "word/".concat(target.replace(/^\//, ''));
                                if (zip.file(zipPath) === null && zip.file(target) !== null) {
                                    zipPath = target;
                                }
                                if (zip.file(zipPath) !== null) {
                                    mimeType = detectMimeType(target);
                                    imageMap.set(relationshipId, {
                                        relationshipId: relationshipId,
                                        targetPath: target,
                                        zipPath: zipPath,
                                        mimeType: mimeType,
                                    });
                                }
                            }
                        }
                    }
                    return [2 /*return*/, __assign(__assign(__assign(__assign({ documentXml: documentXml }, (relationshipsXml !== undefined ? { relationshipsXml: relationshipsXml } : {})), (numberingXml !== undefined ? { numberingXml: numberingXml } : {})), (stylesXml !== undefined ? { stylesXml: stylesXml } : {})), { imageMap: imageMap, zipInstance: zip })];
            }
        });
    });
}
