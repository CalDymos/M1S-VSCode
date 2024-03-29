#!/usr/bin/env node
var start = function start_(options) {

  var m1sindent = require('./indent');
  var fs = require('fs');
  var vscode = require('vscode');
  var options = {};
  var editor = vscode.window.activeTextEditor;


  var document = editor.document;
  var inFile = document.fileName;
  var outFile = inFile;

  for (var i = 0; i < process.argv.length; i++) {
    var arg = process.argv[i];
    var option = arg.replace('--', '');

    switch (option) {
      case 'FormatterLevel':
        var value = process.argv[++i];
        if (!isNaN(value)) {
          options[option] = Number(value);
        } else {
          console.error("Option '" + option + "' is expected to be a number, got '" + value + "' instead");
          process.exit(1);
        }
        break;
      case 'FormatterIndentChar':
        var value = process.argv[++i];
        if (/^(\ +|\\t)$/.test(value)) {
          options[option] = value.replace('\\t', '\t');
        } else {
          console.error("Option '" + option + "' accepts tabs or spaces, got '" + value + "' instead");
          process.exit(2);
        }
        break;
      case 'FormatterBreakLineChar':
        var value = process.argv[++i];
        if (/^(\\n|\\r\\n)$/.test(value)) {
          options[option] = value.replace('\\n', '\n').replace('\\r', '\r');
        } else {
          console.error("Option '" + option + "' accepts \\n or \\r\\n only, got '" + value + "' instead");
          process.exit(3);
        }
        break;
      case 'FormatterBreakOnSeperator':
        options[option] = true;
        break;
      case 'FormatterRemoveComments':
        options[option] = true;
        break;
      case 'FormatterOutput':
        var value = process.argv[++i];
        if (value) {
          outFile = value;
        } else {
          console.error("Option '" + option + "' expects a value.");
          process.exit(4);
        }
        break;
      default:
        console.warn("Option '" + option + "' is not valid, will be ignored.");
    }
  }

  console.info('Formatting m1s file:', inFile);
  var data = document.getText();

  var bsource = m1sindent({
    FormatterLevel: options.FormatterLevel,
    FormatterIndentChar: options.FormatterIndentChar,
    FormatterBreakLineChar: options.FormatterBreakLineChar,
    FormatterBreakOnSeperator: options.FormatterBreakOnSeperator,
    FormatterRemoveComments: options.FormatterRemoveComments,
    source: data
  });

  console.info('Writing to m1s file:', outFile);
  editor.edit(editBuilder => {
    editBuilder.replace(document.validateRange(new vscode.Range(0, 0, document.lineCount, 0)), bsource);
  });
  console.info('Done!');

}
