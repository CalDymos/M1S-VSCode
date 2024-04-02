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
const vscode = __importStar(require("vscode"));
const PATTERNS = __importStar(require("./patterns"));
class M1SColorProvider {
    constructor() {
        this.MAX_COLOR = 0xff;
    }
    provideDocumentColors(doc) {
        const array = new Array();
        let matches;
        while ((matches = PATTERNS.COLOR.exec(doc.getText())) !== null) {
            const pos = doc.positionAt(matches.index);
            const posEnd = doc.positionAt(matches.index + matches[0].length);
            const range = new vscode.Range(pos, posEnd);
            let color;
            if (matches[1]) {
                switch (matches[1].toLowerCase()) {
                    case "vbblack":
                        color = new vscode.Color(0, 0, 0, 1);
                        break;
                    case "vbblue":
                        color = new vscode.Color(0, 0, 1, 1);
                        break;
                    case "vbcyan":
                        color = new vscode.Color(0, 1, 1, 1);
                        break;
                    case "vbgreen":
                        color = new vscode.Color(0, 1, 0, 1);
                        break;
                    case "vbmagenta":
                        color = new vscode.Color(1, 0, 1, 1);
                        break;
                    case "vbred":
                        color = new vscode.Color(1, 0, 0, 1);
                        break;
                    case "vbwhite":
                        color = new vscode.Color(1, 1, 1, 1);
                        break;
                    case "vbyellow":
                        color = new vscode.Color(1, 1, 0, 1);
                        break;
                }
            }
            else if (matches[2]) {
                let r, g, b;
                if (matches[3].toLowerCase().startsWith("&h"))
                    r = Number.parseInt(matches[3].substr(2), 16) / this.MAX_COLOR;
                else
                    r = Number.parseInt(matches[3]) / this.MAX_COLOR;
                if (matches[4].toLowerCase().startsWith("&h"))
                    g = Number.parseInt(matches[4].substr(2), 16) / this.MAX_COLOR;
                else
                    g = Number.parseInt(matches[4]) / this.MAX_COLOR;
                if (matches[5].toLowerCase().startsWith("&h"))
                    b = Number.parseInt(matches[5].substr(2), 16) / this.MAX_COLOR;
                else
                    b = Number.parseInt(matches[5]) / this.MAX_COLOR;
                color = new vscode.Color(r, g, b, 1);
            }
            else if (matches[6] && (/color/i).test(doc.lineAt(pos.line).text)) {
                const r = Number.parseInt(matches[6].substr(2, 2), 16) / this.MAX_COLOR;
                const b = Number.parseInt(matches[6].substr(4, 2), 16) / this.MAX_COLOR;
                const g = Number.parseInt(matches[6].substr(6, 2), 16) / this.MAX_COLOR;
                color = new vscode.Color(r, g, b, 1);
            }
            if (color)
                array.push(new vscode.ColorInformation(range, color));
        }
        return array;
    }
    provideColorPresentations(color) {
        return [new vscode.ColorPresentation(`RGB(${color.red * this.MAX_COLOR}, ${color.green * this.MAX_COLOR}, ${color.blue * this.MAX_COLOR})`)];
    }
}
exports.default = vscode.languages.registerColorProvider({ scheme: "file", language: "m1s" }, new M1SColorProvider());
//# sourceMappingURL=colorprovider.js.map