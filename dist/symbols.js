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
const showVariableSymbols = vscode_1.workspace.getConfiguration("m1s").get("showVariableSymbols");
const showParameterSymbols = vscode_1.workspace.getConfiguration("m1s").get("showParamSymbols");
const FUNCTION = RegExp(PATTERNS.FUNCTION.source, "i");
const CLASS = RegExp(PATTERNS.CLASS.source, "i");
const PROP = RegExp(PATTERNS.PROP.source, "i");
function provideDocumentSymbols(doc) {
    const result = [];
    const varList = [];
    const Blocks = [];
    for (let lineNum = 0; lineNum < doc.lineCount; lineNum++) {
        const line = doc.lineAt(lineNum);
        if (line.isEmptyOrWhitespace || line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
            continue;
        const LineTextwithoutComment = (/^([^'\n\r]*).*$/m).exec(line.text);
        for (const lineText of LineTextwithoutComment[1].split(":")) {
            let name;
            let symbol;
            let matches = [];
            if ((matches = CLASS.exec(lineText)) !== null) {
                name = matches[3];
                symbol = new vscode_1.DocumentSymbol(name, "", vscode_1.SymbolKind.Class, line.range, line.range);
            }
            else if ((matches = FUNCTION.exec(lineText)) !== null) {
                name = matches[4];
                let detail = "";
                let symKind = vscode_1.SymbolKind.Function;
                if (matches[3].toLowerCase() === "sub")
                    if ((/class_(initialize|terminate)/i).test(name)) {
                        symKind = vscode_1.SymbolKind.Constructor;
                    }
                    else {
                        detail = "Sub";
                    }
                else {
                    detail = "Function";
                }
                if (showParameterSymbols)
                    name = matches[5];
                symbol = new vscode_1.DocumentSymbol(name, detail, symKind, line.range, line.range);
                if (showParameterSymbols) {
                    if (matches[6])
                        matches[6].split(",").forEach(param => {
                            symbol.children.push(new vscode_1.DocumentSymbol(param.trim(), "Parameter", vscode_1.SymbolKind.Variable, line.range, line.range));
                        });
                }
            }
            else if ((matches = PROP.exec(lineText)) !== null) {
                name = matches[4];
                symbol = new vscode_1.DocumentSymbol(name, matches[3], vscode_1.SymbolKind.Property, line.range, line.range);
                if ((/Default[\t ]*Property[\t ]*Get/i).test(matches[2]))
                    symbol.detail = "Default Get";
            }
            else if (showVariableSymbols) {
                while ((matches = PATTERNS.VAR.exec(lineText)) !== null) {
                    const varNames = matches[2].split(",");
                    for (const varname of varNames) {
                        const vname = varname.replace(PATTERNS.ARRAYBRACKETS, "").trim();
                        if (varList.indexOf(vname) === -1 || !(/\bSet\b/i).test(matches[0])) {
                            varList.push(vname);
                            let symKind = vscode_1.SymbolKind.Variable;
                            if ((/\bconst\b/i).test(matches[1]))
                                symKind = vscode_1.SymbolKind.Constant;
                            else if ((/\bSet\b/i).test(matches[0]))
                                symKind = vscode_1.SymbolKind.Struct;
                            else if ((/\w+[\t ]*\([\t ]*\d*[\t ]*\)/i).test(varname))
                                symKind = vscode_1.SymbolKind.Array;
                            const r = new vscode_1.Range(line.lineNumber, 0, line.lineNumber, PATTERNS.VAR.lastIndex);
                            const variableSymbol = new vscode_1.DocumentSymbol(vname, "", symKind, r, r);
                            if (Blocks.length === 0)
                                result.push(variableSymbol);
                            else
                                Blocks[Blocks.length - 1].children.push(variableSymbol);
                        }
                    }
                }
            }
            if (symbol) {
                if (Blocks.length === 0)
                    result.push(symbol);
                else
                    Blocks[Blocks.length - 1].children.push(symbol);
                Blocks.push(symbol);
            }
            if ((matches = PATTERNS.ENDLINE.exec(lineText)) !== null)
                Blocks.pop();
        }
    }
    return result;
}
exports.default = vscode_1.languages.registerDocumentSymbolProvider({ scheme: "file", language: "m1s" }, { provideDocumentSymbols });
//# sourceMappingURL=symbols.js.map