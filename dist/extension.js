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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.activate = void 0;
const vscode_1 = require("vscode");
const hover_1 = __importDefault(require("./hover"));
const completion_1 = __importDefault(require("./completion"));
const symbols_1 = __importDefault(require("./symbols"));
const signature_1 = __importDefault(require("./signature"));
const definition_1 = __importDefault(require("./definition"));
const colorprovider_1 = __importDefault(require("./colorprovider"));
const Launcher_1 = __importDefault(require("./Launcher"));
const cmds = __importStar(require("./commands"));
const Includes_1 = require("./Includes");

const m1sindent = require('./indent');
const path = require('path');
const {
	Position,
	Range,
} = require('vscode');

const contributions = vscode_1.workspace.getConfiguration('m1svscode');
const indentCharValue = '\t';
const breakLineCharValue = '\n';
const levelValue = contributions.get('FormatterLevel');
const breakOnSeperatorValue = contributions.get('FormatterBreakOnSeperator');
const removeCommentsValue = contributions.get('FormatterRemoveComments');
const editor = vscode_1.window.activeTextEditor;

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
	//vscode_1.window.showInformationMessage('M1S is now active!');

    Includes_1.Includes.set("Global", new Includes_1.IncludeFile(context.asAbsolutePath("./GlobalDefs.vbs")));
    Includes_1.Includes.set("ObjectDefs", new Includes_1.IncludeFile(context.asAbsolutePath("./ObjectDefs.vbs")));
    vscode_1.workspace.onDidChangeConfiguration(Includes_1.reloadImportDocuments);
    Includes_1.reloadImportDocuments();
    
    vscode_1.commands.registerCommand("m1s.runScript", () => {
        cmds.runScript();
    });
    vscode_1.commands.registerCommand("m1s.killScript", () => {
        cmds.killScript();
    });
    	
    let buttonActivation = vscode_1.commands.registerTextEditorCommand('extension.indenter', (editor) => {
	let document = editor.document
	prepareDocument(document)
    });

    let formatFunction = vscode_1.languages.registerDocumentFormattingEditProvider('m1s', {
	provideDocumentFormattingEdits: (document) => {
		prepareDocument(document)
		}
    });

    context.subscriptions.push(hover_1.default, completion_1.default, symbols_1.default, signature_1.default, definition_1.default, colorprovider_1.default, Launcher_1.default.launchConfigProvider, Launcher_1.default.inlineDebugAdapterFactory, buttonActivation, formatFunction);
}

function prepareDocument(document) {

	if (editor && editor.document.languageId === 'm1s') {

		var inFile = document.fileName;
		const documentText = document.getText();
		const fileExtension = getFileExtension(inFile);
		const start = new Position(getStartLine(documentText, fileExtension), 0);
		const end = new Position(document.lineCount + 1, 0);
		const range = new Range(start, end);
		const sourceFile = document.getText(range);

		vscode_1.window.showInformationMessage('Formatting Script');

		let outFile = m1sindent({
			level: levelValue,
			indentChar: indentCharValue,
			breakLineChar: breakLineCharValue,
			breakOnSeperator: breakOnSeperatorValue,
			removeComments: removeCommentsValue,
			source: sourceFile,
		});
		console.log(outFile);
		const edit = new vscode_1.WorkspaceEdit();
		edit.replace(document.uri, range, outFile);
		return vscode_1.workspace.applyEdit(edit)

	} else {
		vscode_1.window.showInformationMessage('Not a Mach3 Script file!');
	}
}

function getStartLine(text, fileExtension) {

	/*if (fileExtension === '.m1s') {
		const lineNumber = findLineNumber(text, '#expand');
		return lineNumber;
	}*/
	//else {
		return 0;
	//}
}

function findLineNumber(text, searchText) {
	const lines = text.split('\n');
	const lineNumber = lines.findIndex(line => line.includes(searchText)) + 1;
	return lineNumber;
}

function getFileExtension(fileName) {
	return path.extname(fileName);
}

exports.activate = activate;
//# sourceMappingURL=extension.js.map
