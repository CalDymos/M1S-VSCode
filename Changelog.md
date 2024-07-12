# M1S-VSCode Changelog

## 0.20.8

- Output file name during compilation corrected.
- Bugfix in 'vscce' (stderr adjusted for 'syntax error')
- RegEx pattern for 'INCLUDEFILE' adjusted

## 0.20.7

- Manifest for 'vscce' has been removed because of problems on some computers
- change self Registration for 'vscce.exe' (Add Note to README)

## 0.20.6

- prevent change of Class ID between Version changes for 'vscce.exe' 
- added self Registration for ActiveX Exe 'vscce.exe'
- added Manifest for 'vscce.exe'
- Include the dependencies 'msvbvm60.dll' and 'msxml6.dll' in the package

## 0.20.5

- fix bug in const diagnostic
- fix some bugs in 'vscce'

## 0.20.4

-  diagnostics for non-existent include files and non-initialized constants extended
-  fix some bugs

## 0.20.3

- implements "full syntax check"
- diagnistic adjusted and Expand file is opened if there is an error in the expand file. (syntax check)
- 'Go to Definition' for Expand Files added

## 0.20.2

- Shows error line in VS Code if an error occurs during compilation

## 0.20.1

- Context menu “M1S Formatter” removed (duplicate entry for one and the same function)
  Formatting can still be performed via the context menu entry “Format Document” or the shortcut Shift+Alt+F.
- Add context menu entry for "full syntax check" and "Compile Script".
- Remove Commands "killScript" and "runnScript" because of the use of the internal launch config, instead of a user command.
- Added external interpreter (vscce.exe) to compile and debug the script. Currently only compiling is possible.

## 0.14.12

- fix some bugs

## 0.14.11

- RegEx pattern for 'fields' optimized
- fix error in diagnostic
- semantic highlight extended

## 0.14.10

- fix Error in Pattern for If ... Then diagnostic.

## 0.14.9

- add Fields to Outline
- update GlobalDefs
- First implementation of diagnostic informations
- external files are now included in auto-completition (via #expand)

## 0.14.8

- Fixed bug that #Region was not folded.

## 0.14.7

- Add #Region comment for a better outline
- Readme updated

## 0.14.6

- Formatting for constants / parameters with preceding minus adjusted
- Formatting adjusted for short type declaration

## 0.14.5

- Outline corrected / adjusted
- change extension of Import Files (GlobalDefs / ObjectDefs)
- Add syntax highlighting and formatting for 'Declare'

## 0.14.4

- Fixed errors in the formatting of File IO Keywords (Open, close, input etc.)

## 0.14.3

- Fixed errors in the formatting of labels and argument assignments

## 0.14.2

- Fix formatting for minus parameter values
- Highlight gLobal Vars.

## 0.14.1

- Hotfix for formatting problem with #expand

## 0.14.0

- snippets adjusted
- GlobalDefs adjusted
- Syntax highlighting adjusted
- Readme updated

## 0.13.0

- Add Code Formatter based on [VBA Formatter v0.0.3](https://github.com/threatcon/vba-formatter.git)

## 0.12.1

- First pre-release based on [VBSCript Extension for Visual Studio Code v1.21.1](https://github.com/Serpen/VBS-VSCode/releases/tag/1.2.1)
