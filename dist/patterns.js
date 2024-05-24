"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.COLOR = exports.REGION = exports.ENDREGION = exports.TYPE = exports.INCLUDES = exports.ARRAYBRACKETS = exports.ENDLINE = exports.PARAM_SUMMARY = exports.ENDTYPE = exports.FIELD
exports.COMMENT_SUMMARY = exports.DEFVAR = exports.DEF = exports.VAR_COMPLS = exports.VAR = exports.PROP = exports.CLASS = exports.FUNCTION = exports.IF = exports.IFTHEN = void 0;
exports.FUNCTION = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:(?:Public|Private)[\t ]+)?(Function|Sub)[\t ]+((\[?[a-z]\w*\]?)[\t ]*(?:\((.*)\))?))/img;
exports.CLASS = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:(?:Public|Private)[\t ]+)?Class[\t ]+(\[?[a-z]\w*\]?))/img;
exports.PROP = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+(\[?[a-z]\w*\]?))(?:\((.*)\))?/img;
exports.VAR = /(?<!'\s*)(?:^|:)[\t ]*(Dim|Set|Const|Private[\t ]+Const|Public[\t ]+Const|Private|Public|Global)[\t ]+(?!Sub|Function|Class|Property)([a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?(?:[\t ]*,[\t ]*[a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?)*)[\t ]*.*(?:$|:)/img;
exports.FIELD = /^[\t ]*([a-z][a-z0-9_]*)[\t ]*(?:As)[\t ]+([^ \(\'\n\r]*)(?:[\s]*[\n\r]|[ ]+[']+[ ]*([\w\d]*))/img;
exports.VAR_COMPLS = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class|Property [GLT]et))[\t ]+\w+[^:]*$/i;
function DEF(input, word) {
    return new RegExp(`((?:^[\\t ]*'.*$(?:\\r\\n|\\n))*)^[^'\\n\\r]*^[\\t ]*((?:(?:(?:(?:Private[\\t ]+|Public[\\t ]+)?(?:Class|Function|Sub|Property[\\t ][GLS]et)))[\\t ]+)(\\b${word}\\b).*)$`, "im").exec(input);
}
exports.DEF = DEF;
function DEFVAR(input, word) {
    return new RegExp(`((?:^[\\t ]*'.*$(?:\\r\\n|\\n))*)^[^'\\n\\r]*^[\\t ]*((?:(?:Const|Dim|(?:Private|Public|Global)(?![\\t ]+(?:Sub|Function)))[\\t ]+)[\\w\\t ,]*(\\b${word}\\b).*)$`, "im").exec(input);
}
exports.DEFVAR = DEFVAR;
exports.COMMENT_SUMMARY = /(?:'''\s*<summary>|'\s*)([^<\n\r]*)(?:<\/summary>)?/i;
function PARAM_SUMMARY(input, word) {
    return new RegExp(`'''\\s*<param name=["']${word}["']>(.*)<\\/param>`, "i").exec(input);
}
exports.PARAM_SUMMARY = PARAM_SUMMARY;
exports.ENDLINE = (/(?:^|:)[\t ]*End\s+(Sub|Class|Function|Property|Type)/i);
exports.ARRAYBRACKETS = /\(\s*\d*\s*\)/;
exports.COLOR = /\b(RGB[\t ]*\([\t ]*(&h[0-9a-f]+|\d+)[\t ]*,[\t ]*(&h[0-9a-f]+|\d+)[\t ]*,[\t ]*(&h[0-9a-f]+|\d+)[\t ]*\))|(&h[0-9a-f]{6}\b)/ig;
exports.INCLUDES = /#\bexpand\b[\t ]*([<])([^"<.\n\r]+)([>])|(["])([^"<\n\r]+)(["])/ig;
exports.TYPE = /(?:(Type))[\t ]+(\[?[a-zA-Z]\w*\]?)/ig;
exports.ENDTYPE = (/(?:^|:)[\t ]*End\s+(Type)/i);
exports.REGION = /(?:'#Region\s)([^<\n\r]*)/i;
exports.ENDREGION = /(?:'#End Region)/i;
exports.IFTHEN = /^[\t ]*(Elseif|If)[\t ]+[^\n\r]*(?:[\t \)](\bThen\b))/i;
exports.IF = /^[\t ]*(Elseif|If)[\t ]*/i;
//# sourceMappingURL=patterns.js.map