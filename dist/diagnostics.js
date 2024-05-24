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
Object.defineProperty(exports, "__esModule", { value: true });
exports.subscribeToDocumentChanges = void 0;
const vscode_1 = __importStar(require("vscode"));
const PATTERNS = __importStar(require("./patterns"));

function refreshDiagnostics(doc, m1sDiagnostics) {
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
                    diagnostics.push(createDiagnostic(doc, line, lineNum, index, 'missing_then'));
                }
            }
        }
        m1sDiagnostics.set(doc.uri, diagnostics);
    }
}
function createDiagnostic(doc, line, lineNum, startChr, diagCode) {
    const range = new vscode_1.Range(lineNum, startChr, lineNum, line.text.length - startChr);
    const diagnostic = new vscode_1.Diagnostic(range, "Error, If / ElseIf ... Then", vscode_1.DiagnosticSeverity.Error);
    diagnostic.code = diagCode;
    return diagnostic;
}
function subscribeToDocumentChanges(context, m1sDiagnostics) {
    if (vscode_1.window.activeTextEditor) {
        refreshDiagnostics(vscode_1.window.activeTextEditor.document, m1sDiagnostics);
    }
    context.subscriptions.push(vscode_1.window.onDidChangeActiveTextEditor(editor => {
        if (editor) {
            refreshDiagnostics(editor.document, m1sDiagnostics);
        }
    }));
    context.subscriptions.push(vscode_1.workspace.onDidChangeTextDocument(e => refreshDiagnostics(e.document, m1sDiagnostics)));
    context.subscriptions.push(vscode_1.workspace.onDidCloseTextDocument(doc => m1sDiagnostics.delete(doc.uri)));
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