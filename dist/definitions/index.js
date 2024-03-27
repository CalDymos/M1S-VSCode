"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
var _a, _b;
Object.defineProperty(exports, "__esModule", { value: true });
const keywords_json_1 = __importDefault(require("./keywords.json"));
const operators_json_1 = __importDefault(require("./operators.json"));
const vscode_1 = require("vscode");
const completions = new Array();
for (const entry in keywords_json_1.default) {
    const itm = new vscode_1.CompletionItem(entry, vscode_1.CompletionItemKind.Keyword);
    itm.detail = entry;
    itm.documentation = (_a = keywords_json_1.default[entry]) === null || _a === void 0 ? void 0 : _a.documentation;
    completions.push(itm);
}
for (const entry in operators_json_1.default) {
    const itm = new vscode_1.CompletionItem(entry, vscode_1.CompletionItemKind.Operator);
    itm.detail = entry;
    itm.documentation = (_b = operators_json_1.default[entry]) === null || _b === void 0 ? void 0 : _b.documentation;
    itm.filterText = `Operator ${entry}`;
    completions.push(itm);
}
exports.default = completions;
//# sourceMappingURL=index.js.map