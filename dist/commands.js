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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.compileScript = exports.checkScript = void 0;
const vscode_1 = require("vscode");
const childProcess = __importStar(require("child_process"));
const path_1 = __importDefault(require("path"));
const localize_1 = __importDefault(require("./localize"));
const fs = __importStar(require("fs"));
const diagCollection = vscode_1.languages.createDiagnosticCollection("m1s");
const configuration = vscode_1.workspace.getConfiguration("m1s");
const m1sOut = vscode_1.window.createOutputChannel("Mach3Script");
let runner;
const scriptInterpreter = configuration.get("interpreter");
const outputDir = configuration.get("outputFolder");
const mach3Dir = configuration.get("mach3Dir");
const useScreenSet = configuration.get("useScreenSet");
const useProfile = configuration.get("useProfile");
let statbar;

function compileScript() {
    if (!vscode_1.window.activeTextEditor)
        return;
    const doc = vscode_1.window.activeTextEditor.document;
    doc.save().then(() => {
        m1sOut.clear();
        m1sOut.show(true);
        if (statbar)
            statbar.dispose();
        statbar = vscode_1.window.setStatusBarMessage(localize_1.default("m1s.msg.compilescript"));
        let workDir;
        if (vscode_1.workspace.workspaceFolders != null)
            workDir = vscode_1.workspace.workspaceFolders[0].uri.fsPath;
        else
            workDir = path_1.default.dirname(doc.fileName);
        const extDir = vscode_1.extensions.getExtension('caldymos.m1svscode').extensionUri.fsPath
        const srcDir = path_1.default.dirname(doc.fileName);
        let CEinterpreter = scriptInterpreter;
        let outDir = outputDir;

        CEinterpreter = CEinterpreter.replace("${workspaceFolder}", workDir);
        CEinterpreter = CEinterpreter.replace("${extensionFolder}", extDir);
        CEinterpreter = CEinterpreter.replace("${fileDirname}", srcDir);
        outDir = outDir.replace("${workspaceFolder}", workDir);
        outDir = outDir.replace("${extensionFolder}", extDir);
        outDir = outDir.replace("${fileDirname}", srcDir);
        let cmdline = ' -cmd:compile' +
                      ' -src:\'' + doc.fileName + '\'' +
                      ' -mach3Dir:\'' + mach3Dir + '\'' +
                      ' -ScreenSet:\'' + useScreenSet + '\'' +
                      ' -Profile:\'' + useProfile + '\'' +
                      ' -outputFolder:\'' + outDir + '\''
        runner = childProcess.spawn(CEinterpreter, [cmdline], {
            cwd: workDir, windowsHide: true, timout: 10000

        });
        runner.stdout.on("data", data => {
            const output = data.toString();
            m1sOut.append(output);
        });
        runner.stderr.on("data", data => {
            const output = data.toString();
            const match = (/.*Error on line: (\d+) - (.*)/).exec(output);
            if (match) {
                const line = Number.parseInt(match[1]) - 1;
                const diag = new vscode_1.Diagnostic(new vscode_1.Range(line, 0, line, doc.lineAt(line).text.length), match[2], vscode_1.DiagnosticSeverity.Error);
                diagCollection.set(vscode_1.Uri.file(doc.fileName), [diag]);
            }
            m1sOut.append(output);
        });
        runner.on("exit", code => {
            m1sOut.appendLine(`Process exited with code ${code}`);
            statbar.dispose();
        });
        runner.on('error', (err) => {
            console.log(`Error launching command: ${err}`);
        });
    }, () => {
        vscode_1.window.showErrorMessage("Document can't be saved");
        return;
    });
}
exports.compileScript = compileScript;

function checkScript() {
    if (!vscode_1.window.activeTextEditor)
        return;
    const doc = vscode_1.window.activeTextEditor.document;
}
exports.checkScript = checkScript;

function killProcess() {
    runner === null || runner === void 0 ? void 0 : runner.kill();
    statbar === null || statbar === void 0 ? void 0 : statbar.dispose();
}
//# sourceMappingURL=commands.js.map