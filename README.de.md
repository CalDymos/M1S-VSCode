# Mach3 Script Extension für Visual Studio Code
Diese Erweiterung bietet Sprachunterstützung für Mach3 Script für [Visual Studio Code](https://code.visualstudio.com/).

### [English](README.md) | Deutsch


## Features
- Formatierung (über Kontextmenü)

- Gliederung
- Autovervollständigung

![Outline](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/Completion-And-Outline.png)

- Gehe zu Definition

- Hover

![Hover](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/Hover.png)

- Signaturen

![Signature](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/Signature.png)

- Farbinformationen anzeigen

![ColorProvider](https://github.com/CalDymos/M1S-VSCode/raw/master/assets/docs/ColorProvider.png)

- Code Vorlagen

- Higlight von Problemen / Fehlern

- Vollständige Syntaxprüfung über Cypress Enable Compiler

- Kompilieren des Skripts über Cypress Enable Compiler

## Funktionen noch nicht vollständig implementiert (noch in Arbeit)

- Hervorheben und Anzeigen von Problemen / Fehlern

## Wie funktioniert die Erweiterung

#### Formatierung
- Klicken Sie mit der rechten Maustaste in das Codefenster -> wählen Sie den Kontextmenüeintrag "Dokument formatieren".<br>
  Die Einzüge und die Schreibweise der Schlüsselwörter werden im aktuellen Dokument angepasst.<br>
  Optional können auch Kommentare entfernt und Zeilen mit ':' umbrochen werden (siehe Einstellungen)

#### Vervollständigung
- Bei der Autovervollständigung wird der aktuelle Quellcode, alle Include-Dateien (#expand),<br>
  die Datei "objectDefs.m1s" und die Datei "GlobalDefs.m1s" nach Definitionen durchsucht.

#### Gehe zu Definition
- Klicken Sie mit der rechten Maustaste auf eine Variable,<br> Prozedur oder Include-Datei und wählen Sie den Kontextmenüeintrag 'Gehe zur Definition'.

#### Vollständige Syntaxprüfung
- Klicken Sie mit der rechten Maustaste in das Codefenster und wählen Sie den Kontextmenüeintrag 'full syntax check'.<br>
  Die Einstellungen 'Mach3Dir' und 'UseScreenSet' müssen korrekt gesetzt sein.<br>
  Am besten legen Sie im jeweiligen Arbeitsbereich (Ordner) einen neuen Ordner '.vscode' an und<br> 
  setzen die Einstellungen in diesem Ordner über die Datei 'settings.json',<br> sodass diese nur für diesen Arbeitsbereich gültig sind.

  **Wenn die Fehlermeldung "Unerwarteter Fehler, Beenden" erscheint, <br>muss die Datei 'vscce.exe' einmal mit Administratorrechten ausgeführt werden, damit die COM-Schnittstellen korrekt registriert werden.**

#### Skript compilieren
- Klicken Sie mit der rechten Maustaste in das Codefenster und wählen Sie den Kontextmenüeintrag 'Skript kompilieren'.<br>
  Die Einstellungen 'Mach3Dir', 'UseScreenSet' und 'OutputFolder' müssen korrekt gesetzt sein.<br>
  Am besten legen Sie im jeweiligen Arbeitsbereich (Ordner) einen neuen Ordner '.vscode' an und<br> 
  setzen die Einstellungen in diesem Ordner über die Datei 'settings.json',<br> sodass diese nur für diesen Arbeitsbereich gültig sind.

  **Wenn die Fehlermeldung "Unerwarteter Fehler, Beenden" erscheint, <br>muss die Datei 'vscce.exe' einmal mit Administratorrechten ausgeführt werden, damit die COM-Schnittstellen korrekt registriert werden.**
  

## Mitarbeit
Du kannst dieses Projekt unterstützen, indem du die Quelldateien forkst und einen Pull Request/PR mit deinen Veränderungen erzeugst oder eine Issue mit deinem Problem/deiner Idee erzeugst.
- Vervollständigung der M1S Sprachdokumentation
- Übersetzung in weitere Sprachen
- ...


## Referenzen / Danksagung
Diese Erweiterung basiert auf VBScript Extension für Visual Studio Code von serpen.

