"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function() { return m[k]; } });
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const vscode_1 = require("vscode");
const PATTERNS = __importStar(require("./patterns"));
const Includes_1 = require("./Includes");
function GetHover(docText, lookup, scope) {
    const results = [];
    let matches = PATTERNS.DEF(docText, lookup);
    if (matches) {
        if (matches[1]) {
            const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
            if (summary[1])
                results.push(new vscode_1.Hover({ language: "m1s", value: `${matches[2]} ' [${scope}]\n' ${summary[1]}` }));
            else
                results.push(new vscode_1.Hover({ language: "m1s", value: `${matches[2]} ' [${scope}]` }));
        }
        else
            results.push(new vscode_1.Hover({ language: "m1s", value: `${matches[2]} ' [${scope}]` }));
    }
    matches = PATTERNS.DEFVAR(docText, lookup);
    if (matches) {
        if (matches[1]) {
            const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
            if (summary[1])
                results.push(new vscode_1.Hover({ language: "m1s", value: `${matches[2]} ' [${scope}]\n' ${summary[1]}` }));
            else
                results.push(new vscode_1.Hover({ language: "m1s", value: `${matches[2]} ' [${scope}]` }));
        }
        else
            results.push(new vscode_1.Hover({ language: "m1s", value: `${matches[2]} ' [${scope}]` }));
    }
    return results;
}
function GetParamHover(text, lookup) {
    var _a;
    const hovers = [];
    let matches;
    while (matches = PATTERNS.FUNCTION.exec(text))
        (_a = matches[6]) === null || _a === void 0 ? void 0 : _a.split(",").filter(p => p.trim() === lookup).forEach(() => {
            hovers.push(new vscode_1.Hover({ language: "m1s", value: `${lookup} ' [Parameter]` }));
        });
    if (hovers.length > 0)
        return [hovers[hovers.length - 1]];
    else
        return [];
}
function provideHover(doc, position) {
    const wordRange = doc.getWordRangeAtPosition(position);
    const word = wordRange ? doc.getText(wordRange) : "";
    const line = doc.lineAt(position).text;
    const hoverresults = [];
    if (word.trim() === "")
        return null;
    if (!new RegExp(`^[^']*${word}`).test(line))
        return null;
    let count = 0;
    for (let i = 0; i < position.character; i++)
        if (line[i] === '"')
            count++;
    if (count % 2 === 1)
        return null;
    hoverresults.push(...GetHover(doc.getText(), word, "Local"));
    for (const ExtraDocText of Includes_1.getImportsWithLocal(doc))
        hoverresults.push(...GetHover(ExtraDocText[1].Content, word, ExtraDocText[0]));
    hoverresults.push(...GetParamHover(doc.getText(new vscode_1.Range(new vscode_1.Position(0, 0), new vscode_1.Position(position.line + 1, 0))), word));
    if (hoverresults.length > 0)
        return hoverresults[0];
    else
        return null;
}
exports.default = vscode_1.languages.registerHoverProvider({ scheme: "file", language: "m1s" }, { provideHover });
//# sourceMappingURL=hover.js.map