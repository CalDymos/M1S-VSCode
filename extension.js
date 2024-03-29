const vscode = require('vscode');
const m1sindent = require('./dist/indent');
const path = require('path');
const {
	Position,
	Range,
} = require('vscode');

const contributions = vscode.workspace.getConfiguration('m1svscode');
const indentCharValue = '\t';
const breakLineCharValue = '\n';
const levelValue = contributions.get('level');
const breakOnSeperatorValue = contributions.get('breakOnSeperator');
const removeCommentsValue = contributions.get('removeComments');
const editor = vscode.window.activeTextEditor;

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
	vscode.window.showInformationMessage('Mach3 Script Formatter is now active!');

	let buttonActivation = vscode.commands.registerTextEditorCommand('extension.indent', (editor) => {
		let document = editor.document
		prepareDocument(document)
	});

	context.subscriptions.push(buttonActivation);

	let formatFunction = vscode.languages.registerDocumentFormattingEditProvider('m1s', {
		provideDocumentFormattingEdits: (document) => {
			prepareDocument(document)
		}
	});

	context.subscriptions.push(formatFunction);
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

		vscode.window.showInformationMessage('Formatting Script');

		let outFile = m1sindent({
			level: levelValue,
			indentChar: indentCharValue,
			breakLineChar: breakLineCharValue,
			breakOnSeperator: breakOnSeperatorValue,
			removeComments: removeCommentsValue,
			source: sourceFile,
		});
		console.log(outFile);
		const edit = new vscode.WorkspaceEdit();
		edit.replace(document.uri, range, outFile);
		return vscode.workspace.applyEdit(edit)

	} else {
		vscode.window.showInformationMessage('Not a Mach3 Script file!');
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


function deactivate() {
	vscode.window.showInformationMessage('deactivated');
}

module.exports = {
	activate,
	deactivate
}
