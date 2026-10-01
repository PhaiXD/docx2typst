"use strict";
/**
 * @file errors.ts
 * Custom error classes for DOCX reading and parsing operations.
 */
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.DocxParseError = exports.DocxReadError = void 0;
/**
 * Error thrown when reading, opening, or extracting a .docx file fails.
 */
var DocxReadError = /** @class */ (function (_super) {
    __extends(DocxReadError, _super);
    /**
     * Creates a new instance of DocxReadError.
     *
     * @param message - Human-readable error description.
     * @param options - Optional configuration including the underlying error cause.
     */
    function DocxReadError(message, options) {
        var _newTarget = this.constructor;
        var _this = _super.call(this, message) || this;
        _this.name = 'DocxReadError';
        _this.cause = options === null || options === void 0 ? void 0 : options.cause;
        // Restore prototype chain for instanceof checks across transpilations
        Object.setPrototypeOf(_this, _newTarget.prototype);
        return _this;
    }
    return DocxReadError;
}(Error));
exports.DocxReadError = DocxReadError;
/**
 * Error thrown when parsing OOXML markup or extracting document structures fails.
 */
var DocxParseError = /** @class */ (function (_super) {
    __extends(DocxParseError, _super);
    /**
     * Creates a new instance of DocxParseError.
     *
     * @param message - Human-readable error description.
     * @param options - Optional configuration including the underlying error cause.
     */
    function DocxParseError(message, options) {
        var _newTarget = this.constructor;
        var _this = _super.call(this, message) || this;
        _this.name = 'DocxParseError';
        _this.cause = options === null || options === void 0 ? void 0 : options.cause;
        // Restore prototype chain for instanceof checks across transpilations
        Object.setPrototypeOf(_this, _newTarget.prototype);
        return _this;
    }
    return DocxParseError;
}(Error));
exports.DocxParseError = DocxParseError;
