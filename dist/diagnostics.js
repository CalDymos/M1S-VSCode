"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function (o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
        desc = { enumerable: true, get: function () { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.subscribeToDocumentChanges = void 0;
const vscode_1 = __importStar(require("vscode"));
const PATTERNS = __importStar(require("./patterns"));
const fs = __importStar(require("fs"));
const localize_1 = __importDefault(require("./localize"));
const configuration = vscode_1.workspace.getConfiguration("m1s");
const mach3Dir = configuration.get("mach3Dir");
const useScreenSet = configuration.get("useScreenSet");
exports.diagCollection = vscode_1.languages.createDiagnosticCollection("m1s");
exports.diagCollectionLaunch = vscode_1.languages.createDiagnosticCollection("m1sLaunch");

function refreshDiagnostics(doc, diagCollection) {
    if (doc.languageId === 'm1s') {
        const diagnostics = [];
        for (let lineNum = 0; lineNum < doc.lineCount; lineNum++) {
            const line = doc.lineAt(lineNum);
            let matches = [];
            if (line.isEmptyOrWhitespace || line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
                continue;
            if ((matches = PATTERNS.IF.exec(line.text)) !== null && lineNum != currentLine()) {
                let matches2 = [];
                if ((matches2 = PATTERNS.IFTHEN.exec(line.text)) === null) {
                    const index = line.text.indexOf(matches[1]);
                    diagnostics.push(createDiagnostic(line.text, lineNum, 0, localize_1.default("m1s.error.missingThen"), 1));
                }
            }
            if ((matches = PATTERNS.INCLUDES.exec(line.text)) !== null) {
                let file = matches[2];
                if (matches[1] === '<')
                    file = mach3Dir + "\\ScreenSetMacros\\" + useScreenSet + "\\" + file + ".m1s";
                else if (matches[1] !== '"')
                    file = "";

                if (file !== "") {
                    if (fs.existsSync(file) !== true) {
                        let file = matches[2];
                        let index = line.text.indexOf(file);
                        diagnostics.push(createDiagnostic(line.text, lineNum, 0, localize_1.default("m1s.error.openIncludeFile").replace('${1}', file), 1696));
                    }
                }
            }
            if ((matches = PATTERNS.CONST.exec(line.text)) !== null) {
                let constName = matches[2];
                let matches2 = [];
                if ((matches2 = PATTERNS.INIT_CONST.exec(line.text)) === null || matches2[1].length === 0) {
                let index = line.text.indexOf(constName);
                diagnostics.push(createDiagnostic(line.text, lineNum, 0, localize_1.default("m1s.error.constantReqInit").replace('${1}', constName), 257));
                }
            }
        }
        diagCollection.set(doc.uri, diagnostics);
    }
}
function createDiagnostic(lineText, lineNum, startChr, diagMsg, diagCode) {
    const range = new vscode_1.Range(lineNum, startChr, lineNum, lineText.length - startChr);
    const diagnostic = new vscode_1.Diagnostic(range, diagMsg, vscode_1.DiagnosticSeverity.Error);
    diagnostic.code = diagCode;
    return diagnostic;
}

function subscribeToDocumentChanges(context) {
    if (vscode_1.window.activeTextEditor) {
        refreshDiagnostics(vscode_1.window.activeTextEditor.document, exports.diagCollection);
    }
    context.subscriptions.push(vscode_1.window.onDidChangeActiveTextEditor(editor => {
        if (editor) {
            refreshDiagnostics(editor.document, exports.diagCollection);
        }
    }));
    context.subscriptions.push(vscode_1.workspace.onDidChangeTextDocument(e => refreshDiagnostics(e.document, exports.diagCollection)));
    context.subscriptions.push(vscode_1.workspace.onDidCloseTextDocument(doc => exports.diagCollection.delete(doc.uri)));
}


function currentLine() {
    const activeEditor = vscode_1.window.activeTextEditor;
    if (activeEditor != null) {
        let linenum = activeEditor.selection.active.line;
        return activeEditor.selection.active.line;
    }
}

exports.subscribeToDocumentChanges = subscribeToDocumentChanges;
//# sourceMappingURL=diagnostics.js.map