"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function (o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function () { return m[k]; } });
}) : (function (o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function (o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function (o, v) {
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
exports.getImportsWithLocal = exports.reloadImportDocuments = exports.Includes = exports.IncludeFile = void 0;
const vscode_1 = require("vscode");
const pathns = __importStar(require("path"));
const PATTERNS = __importStar(require("./patterns"));
//const INCLUDES = RegExp(PATTERNS.INCLUDES.source);
const fs = __importStar(require("fs"));
const mach3Dir = vscode_1.workspace.getConfiguration("m1s").get("mach3Dir");
const useScreenSet = vscode_1.workspace.getConfiguration("m1s").get("useScreenSet");
const diagnostics_1 = require("./diagnostics");

class IncludeFile {
    constructor(path) {
        this.Content = "";
        let path2 = path;
        if (!pathns.isAbsolute(path2))
            path2 = pathns.join(vscode_1.workspace.workspaceFolders[0].uri.fsPath, path2);
        this.Uri = vscode_1.Uri.file(path2);
        if (fs.existsSync(path2) && fs.statSync(path2).isFile())
            this.Content = fs.readFileSync(path2).toString();
    }
}
exports.IncludeFile = IncludeFile;
exports.Includes = new Map();
function reloadImportDocuments() {
    for (const key of exports.Includes.keys()) {
        if (key.startsWith("Include"))
            exports.Includes.delete(key);
    }
}
exports.reloadImportDocuments = reloadImportDocuments;

function getImportsWithLocal(doc) {
    var _a;
    const localIncludes = [...exports.Includes];
    const processedMatches = Array();
    let match;

    while ((match = PATTERNS.INCLUDES.exec(doc.getText())) !== null) {
        let file;
        let bracket;
        if (match[2] != null) {
            bracket = match[1];
            file = match[2];
        }
        else {
            bracket = match[4];
            file = match[5];
        }
        if (processedMatches.indexOf(file.toLowerCase()) === -1) {
            let path = mach3Dir;

            if (bracket === '<') {
                path = mach3Dir + "\\ScreenSetMacros\\" + useScreenSet + "\\" + file + ".m1s";
            } else {
                path = file;
            }

            if (fs.existsSync(path) && ((_a = fs.statSync(path)) === null || _a === void 0 ? void 0 : _a.isFile()))
                localIncludes.push([
                    `Include Statement ${file}`,
                    new IncludeFile(path)
                ]);
            else
                {
                //const diag = new vscode_1.Diagnostic(new vscode_1.Range(line, 0, line, doc.lineAt(line).text.length), match[2], vscode_1.DiagnosticSeverity.Error);
                //diagnostics_1.diagCollectionExpand.set(vscode_1.Uri.file(doc.fileName), [diag]);
                processedMatches.push(file.toLowerCase());
                }
        }
    }
    return localIncludes;
}
exports.getImportsWithLocal = getImportsWithLocal;
//# sourceMappingURL=Includes.js.map