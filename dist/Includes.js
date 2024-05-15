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
//exports.getImportsWithLocal = exports.reloadImportDocuments = exports.customIncludePattern = exports.customIncludeDirs = exports.Includes = exports.IncludeFile = void 0;
exports.getImportsWithLocal = exports.reloadImportDocuments = exports.Includes = exports.IncludeFile = void 0;
const vscode_1 = require("vscode");
const pathns = __importStar(require("path"));
const PATTERNS = __importStar(require("./patterns"));
const INCLUDES = RegExp(PATTERNS.INCLUDES.source);
const fs = __importStar(require("fs"));
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
    //const custumIncludePatternStr = vscode_1.workspace.getConfiguration("m1s").get("custumIncludePattern");
    //const customIncludePatternStr = vscode_1.workspace.getConfiguration("m1s").get("customIncludePattern");
    exports.customIncludeDirs = vscode_1.workspace.getConfiguration("m1s").get("customIncludeDirs");
    // if (custumIncludePatternStr !== customIncludePatternStr && custumIncludePatternStr !== "")
    //     exports.customIncludePattern = new RegExp(custumIncludePatternStr, "ig");
    // else
    //     exports.customIncludePattern = new RegExp(customIncludePatternStr, "ig");
    //const SourceImportFiles = vscode_1.workspace.getConfiguration("m1s").get("includes");
    for (const key of exports.Includes.keys()) {
        if (key.startsWith("Include"))
            exports.Includes.delete(key);
    }
    // SourceImportFiles === null || SourceImportFiles === void 0 ? void 0 : SourceImportFiles.forEach((file) => {
    //     exports.Includes.set(`Include ${pathns.basename(file)}`, new IncludeFile(file));
    // });
}
exports.reloadImportDocuments = reloadImportDocuments;
function getImportsWithLocal(doc) {
    var _a;
    const localIncludes = [...exports.Includes];
    const processedMatches = Array();
    let match;
    //while ((match = exports.customIncludePattern.exec(doc.getText())) !== null) {
        while ((match = INCLUDES.exec(doc.getText())) !== null) {
        if (processedMatches.indexOf(match[1].toLowerCase()) === -1) {
            for (const incDir of exports.customIncludeDirs) {
                let incDirResolved = incDir;
                if (incDirResolved === ".")
                    incDirResolved = pathns.dirname(doc.uri.fsPath);
                else if (incDirResolved === "..")
                    incDirResolved = pathns.dirname(pathns.dirname(doc.uri.fsPath));
                else if (incDirResolved === "${workspaceFolder}")
                    if (vscode_1.workspace.workspaceFolders)
                        incDirResolved = vscode_1.workspace.workspaceFolders[0].uri.fsPath;
                const path = pathns.resolve(incDirResolved, match[1]);
                if (fs.existsSync(path) && ((_a = fs.statSync(path)) === null || _a === void 0 ? void 0 : _a.isFile()))
                    localIncludes.push([
                        `Include Statement ${match[1]}`,
                        new IncludeFile(path)
                    ]);
            }
            processedMatches.push(match[1].toLowerCase());
        }
    }
    return localIncludes;
}
exports.getImportsWithLocal = getImportsWithLocal;
//# sourceMappingURL=Includes.js.map