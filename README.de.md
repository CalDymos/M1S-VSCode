# Mach3 Script Extension für Visual Studio Code
Diese Erweiterung bietet Sprachunterstützung für Mach3 Script für [Visual Studio Code](https://code.visualstudio.com/).

<p align="center">
  <a href="./README.md">English</a> | 
  <span>Deutsch</span>
</p>

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

- Zusätzliche M1S Funktions Bibliotheken als M1S Dateien einbinden
```
{ // settings.json
    "m1s.includes": ["c:\\mylibrary.m1s"]
}
```

## Funktionen noch nicht vollständig implementiert (noch in Arbeit)

- vollständige Syntaxprüfung (Ausführen / Prüfen des Makros über den Cypress Enbable Compiler)
- Zusätzliche Dateien (#expand) in die Autovervollständigung einbeziehen.

## Mitarbeit
Du kannst dieses Projekt unterstützen, indem du die Quelldateien forkst und einen Pull Request/PR mit deinen Veränderungen erzeugst oder eine Issue mit deinem Problem/deiner Idee erzeugst.
- Vervollständigung der M1S Sprachdokumentation #21
- Übersetzung in weitere Sprachen
- ...


## Referenzen / Danksagung
Diese Erweiterung basiert auf VBScript Extension für Visual Studio Code von serpen.

