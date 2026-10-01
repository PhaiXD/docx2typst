"use strict";
/**
 * @file xmlParser.ts
 * XML parsing utility configuring fast-xml-parser for OOXML document parsing.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseXml = parseXml;
exports.parseXmlPreserveOrder = parseXmlPreserveOrder;
var fast_xml_parser_1 = require("fast-xml-parser");
var errors_js_1 = require("../errors.js");
/**
 * Shared instance of XMLParser configured for OOXML schemas.
 */
var parser = new fast_xml_parser_1.XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    parseAttributeValue: true,
    trimValues: false,
    tagValueProcessor: function (tagName, tagValue) {
        if (tagName === 'w:t' || tagName === 't') {
            return tagValue;
        }
        return tagValue.trim();
    },
});
/**
 * Parses an XML string into a structured JavaScript object representation.
 *
 * @param xmlString - The raw XML content to parse.
 * @returns Parsed XML tree as a generic key-value object.
 * @throws {DocxParseError} If the XML string cannot be parsed or produces invalid output.
 */
function parseXml(xmlString) {
    if (typeof xmlString !== 'string' || xmlString.trim().length === 0) {
        throw new errors_js_1.DocxParseError('Invalid XML input: expected a non-empty string.');
    }
    var validation = fast_xml_parser_1.XMLValidator.validate(xmlString);
    if (validation !== true) {
        var errorMsg = typeof validation === 'object' && validation.err
            ? validation.err.msg
            : 'Invalid XML structure';
        throw new errors_js_1.DocxParseError("XML validation error: ".concat(errorMsg));
    }
    try {
        var result = parser.parse(xmlString);
        if (result === null || typeof result !== 'object') {
            throw new errors_js_1.DocxParseError('Failed to parse XML: output is not a valid object.');
        }
        return result;
    }
    catch (error) {
        if (error instanceof errors_js_1.DocxParseError) {
            throw error;
        }
        throw new errors_js_1.DocxParseError('Error parsing XML content', { cause: error });
    }
}
/**
 * Shared instance of XMLParser configured with preserveOrder: true for document-order body traversal.
 */
var orderedParser = new fast_xml_parser_1.XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    parseAttributeValue: true,
    trimValues: false,
    tagValueProcessor: function (tagName, tagValue) {
        if (tagName === 'w:t' || tagName === 't') {
            return tagValue;
        }
        return tagValue.trim();
    },
});
/**
 * Parses an XML string into an ordered array of element nodes preserving sibling document order.
 *
 * @param xmlString - The raw XML content to parse.
 * @returns Parsed XML tree as an array of ordered element objects.
 * @throws {DocxParseError} If the XML string cannot be parsed or produces invalid output.
 */
function parseXmlPreserveOrder(xmlString) {
    if (typeof xmlString !== 'string' || xmlString.trim().length === 0) {
        throw new errors_js_1.DocxParseError('Invalid XML input: expected a non-empty string.');
    }
    var validation = fast_xml_parser_1.XMLValidator.validate(xmlString);
    if (validation !== true) {
        var errorMsg = typeof validation === 'object' && validation.err
            ? validation.err.msg
            : 'Invalid XML structure';
        throw new errors_js_1.DocxParseError("XML validation error: ".concat(errorMsg));
    }
    try {
        var result = orderedParser.parse(xmlString);
        if (!Array.isArray(result)) {
            throw new errors_js_1.DocxParseError('Failed to parse XML: output is not a valid array.');
        }
        return result;
    }
    catch (error) {
        if (error instanceof errors_js_1.DocxParseError) {
            throw error;
        }
        throw new errors_js_1.DocxParseError('Error parsing XML content', { cause: error });
    }
}
