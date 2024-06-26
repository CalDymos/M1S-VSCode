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
exports.InlineDebugAdapterFactory = exports.DebugConfigurationProvider = exports.M1sDebugSession = void 0;
const vscode_debugadapter_1 = require("vscode-debugadapter");
const vscode = __importStar(require("vscode"));
const path_1 = require("path");
const child_process_1 = require("child_process");
const configuration = vscode.workspace.getConfiguration("m1s");
const outputDir = configuration.get("outputFolder");
const mach3Dir = configuration.get("mach3Dir");
const useScreenSet = configuration.get("useScreenSet");
const useProfile = configuration.get("useProfile");
const diagnostics_1 = require("./diagnostics");

class M1sDebugSession extends vscode_debugadapter_1.LoggingDebugSession {
    constructor() {
        super();
    }
    initializeRequest(response) {
        response.body = response.body || {};
        response.body.supportsConfigurationDoneRequest = true;
        response.body.supportsCancelRequest = false;
        response.body.supportsTerminateRequest = true;
        this.sendResponse(response);
    }
    async launchRequest(response, args) {
        diagnostics_1.diagCollectionLaunch.clear();
        this.sendResponse(response);
        const configuration = vscode.workspace.getConfiguration("m1s");
        const scriptInterpreter = configuration.get("interpreter");
        let workDir; 
        if (vscode.workspace.workspaceFolders != null)
            workDir = vscode.workspace.workspaceFolders[0].uri.fsPath;
        else
            workDir = path_1.dirname(args.program);
        const extDir = vscode.extensions.getExtension('caldymos.m1svscode').extensionUri.fsPath
        const srcDir = path_1.dirname(args.program);

        let CEinterpreter = scriptInterpreter;        
        let outDir = outputDir;

        CEinterpreter = CEinterpreter.replace("${workspaceFolder}", workDir);
        CEinterpreter = CEinterpreter.replace("${extensionFolder}", extDir);
        CEinterpreter = CEinterpreter.replace("${fileDirname}", srcDir);
        outDir = outDir.replace("${workspaceFolder}", workDir);
        outDir = outDir.replace("${extensionFolder}", extDir);
        outDir = outDir.replace("${fileDirname}", srcDir);
        let cmdline = ' -cmd:run' +
                      ' -src:\'' + args.program + '\'' +
                      ' -mach3Dir:\'' + mach3Dir + '\'' +
                      ' -ScreenSet:\'' + useScreenSet + '\'' +
                      ' -Profile:\'' + useProfile + '\'' +
                      ' -outputFolder:\'' + outDir + '\''
        this._runner = child_process_1.spawn(CEinterpreter, ['-src:' + args.program, '-cmd:run'], { "cwd": workDir });
        this._runner.stdout.on("data", data => {
            this.sendEvent(new vscode_debugadapter_1.OutputEvent(data.toString(), "stdout"));
        });
        this._runner.stderr.on("data", data => {
            const output = data.toString();
            const match = (/.*\((\d+), (\d+)\) (.*)/).exec(output);
            if (match) {
                const line = Number.parseInt(match[1]) - 1;
                const char = Number.parseInt(match[2]) - 1;

                const diag = new vscode.Diagnostic(new vscode.Range(line, char, line, char), match[3], vscode.DiagnosticSeverity.Error);
                diagnostics_1.diagCollectionLaunch.set(vscode.Uri.file(args.program), [diag]);
            }
            this.sendEvent(new vscode_debugadapter_1.OutputEvent(output, "stderr"));
        });
        this._runner.on("exit", code => {
            this.sendEvent(new vscode_debugadapter_1.OutputEvent(`Process [${this._runner.pid}] exited with code ${code} (&H${code.toString(16).toUpperCase()})`));
            this.sendEvent(new vscode_debugadapter_1.TerminatedEvent());
        });
        this._runner.on('error', (err) => {
            console.log(`Error launching debugger: ${err}`);
        });
    }
    terminateRequest(response, _args, _req) {
        var _a;
        (_a = this._runner) === null || _a === void 0 ? void 0 : _a.kill();
        this.sendEvent(new vscode_debugadapter_1.TerminatedEvent());
        this.sendResponse(response);
    }
}
exports.M1sDebugSession = M1sDebugSession;
class DebugConfigurationProvider {
    resolveDebugConfiguration(_folder, config) {
        const editor = vscode.window.activeTextEditor;
        if (!config || !config.request) {
            if (editor && editor.document.languageId === "m1s") {
                config.type = "m1s";
                config.name = "VSCCE";
                config.request = "launch";
                config.program = editor.document.uri.fsPath;
            }
            else
                return;
        }
        if (!config.program || !editor || editor.document.languageId !== "m1s") {
            return vscode.window.showInformationMessage("Cannot find a program to debug").then(() => {
                return undefined;
            });
        }
        return config;
    }
}
exports.DebugConfigurationProvider = DebugConfigurationProvider;
class InlineDebugAdapterFactory {
    createDebugAdapterDescriptor() {
        return new vscode.DebugAdapterInlineImplementation(new M1sDebugSession());
    }
}
exports.InlineDebugAdapterFactory = InlineDebugAdapterFactory;
const launchConfigProvider = vscode.debug.registerDebugConfigurationProvider("m1s", new DebugConfigurationProvider());
const inlineDebugAdapterFactory = vscode.debug.registerDebugAdapterDescriptorFactory("m1s", new InlineDebugAdapterFactory());
exports.default = { launchConfigProvider, inlineDebugAdapterFactory };
//# sourceMappingURL=Launcher.js.map