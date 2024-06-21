# Mach3 Script Extension for Visual Studio Code
This extension implements basic language features of Mach3 Script for [Visual Studio Code](https://code.visualstudio.com/).

### English | [Deutsch](README.de.md)

## Features
- Formatting (via context menu)

- Outline
- Completion
![Outline](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/Completion-And-Outline.png)

- Goto Definition

- Hover 
![Hover](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/Hover.png)

- Signatures
![Signature](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/Signature.png)

- Color Information
![ColorProvider](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/ColorProvider.png)

- Code snippets

- Higlight of problems / errors

- full syntax check via Cypress Enable compiler
  
- compile script via cypress Enable Compiler

## Functions not yet fully implemented (still in progress)

- Higlight of problems / errors (diagnostic)

## How it works

#### Formatting
- Right-click in the code window -> select the context menu entry “Format Document”.<br>
  The indents and spelling of keywords are adjusted in the current document.<br>
  Optionally, comments can also be removed and lines can be wrapped with ':' (see Settings)

#### Completion
- For auto-completion, the current code, all include files (#expand),<br>
  the 'objectDefs.m1s' file and the 'GlobalDefs.m1s' file are searched for definitions.

#### Goto Definition
- Right click on a variable, procedure or include file, select the context menu entry 'Go To Definition'

#### Full syntax check
- Right-click in the code window and select the context menu entry 'full syntax check'.<br>
  The settings 'Mach3Dir' and 'UseScreenSet' must be set correctly.<br>
  It is best to create a new folder '.vscode' in the respective workspace (folder) and set the settings in<br>
  this folder via the 'settings.json' file so that they are only valid for this workspace

  **If the error message "Unexpected error, quitting" appears, <br>the file 'vscce.exe' must be executed once with administrator rights so that the COM interfaces are registered correctly.**

#### Compile script
- Right-click in the code window and select the context menu entry 'compile Script'.<br>
  The settings 'Mach3Dir', 'UseScreenSet' and 'OutputFolder' must be set correctly.<br>
  It is best to create a new folder '.vscode' in the respective workspace (folder) and set the settings in<br>
  this folder via the 'settings.json' file so that they are only valid for this workspace

  **If the error message "Unexpected error, quitting" appears, <br>the file 'vscce.exe' must be executed once with administrator rights so that the COM interfaces are registered correctly.**
  
## Contribute
You can support this project through PR with your changes or simply add an issue with your idea/bug.
- Complete Language Source Documentation
- Translate
- ...

## References / Thanks
This extension is based on the "VBScript Extension for Visual Studio Code" from serpen.
