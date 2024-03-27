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
const Includes_1 = require("./Includes");
const PATTERNS = __importStar(require("./patterns"));
function findExtDef(docText, lookup, docuri) {
    const posloc = [];
    let match = PATTERNS.DEF(docText, lookup);
    if (match) {
        const pos = match.index + match[1].length;
        const line = docText.slice(0, pos).match(/\n/g).length;
        posloc.push(new vscode_1.Location(docuri, new vscode_1.Position(line, 0)));
    }
    match = PATTERNS.DEFVAR(docText, lookup);
    if (match) {
        const line = docText.slice(0, match.index).match(/\n/g).length;
        posloc.push(new vscode_1.Location(docuri, new vscode_1.Position(line, 0)));
    }
    return posloc;
}
function GetParamDef(docText, lookup, thisUri) {
    var _a;
    const locs = [];
    let matches;
    while (matches = PATTERNS.FUNCTION.exec(docText))
        (_a = matches[6]) === null || _a === void 0 ? void 0 : _a.split(",").filter(p => p.trim() === lookup).forEach(() => {
            const line = docText.slice(0, matches.index).match(/\n/g).length;
            locs.push(new vscode_1.Location(thisUri, new vscode_1.Position(line, 0)));
        });
    if (locs.length > 0)
        return [locs[locs.length - 1]];
    else
        return [];
}
function provideDefinition(doc, position) {
    const lookupRange = doc.getWordRangeAtPosition(position);
    const lookup = doc.getText(lookupRange);
    const docText = doc.getText();
    const posLoc = [];
    let match = PATTERNS.DEF(docText, lookup);
    if (match)
        posLoc.push(new vscode_1.Location(doc.uri, doc.positionAt(match.index)));
    match = PATTERNS.DEFVAR(docText, lookup);
    if (match)
        posLoc.push(new vscode_1.Location(doc.uri, doc.positionAt(match.index)));
    for (const item of Includes_1.getImportsWithLocal(doc))
        posLoc.push(...findExtDef(item[1].Content, lookup, item[1].Uri));
    posLoc.push(...GetParamDef(doc.getText(new vscode_1.Range(new vscode_1.Position(0, 0), new vscode_1.Position(position.line + 1, 0))), lookup, doc.uri));
    return posLoc;
}
exports.default = vscode_1.languages.registerDefinitionProvider({ scheme: "file", language: "m1s" }, { provideDefinition });
//# sourceMappingURL=definition.js.map