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
const vscode_1 = require("vscode");
const definitions_1 = __importDefault(require("./definitions"));
const Includes_1 = require("./Includes");
const PATTERNS = __importStar(require("./patterns"));
const ObjectSourceImportName = "ObjectDefs";
function getVariableCompletions(text, scope) {
    const CIs = [];
    const foundVals = new Array();
    let matches;
    while (matches = PATTERNS.VAR.exec(text)) {
        for (const match of matches[2].split(",")) {
            const name = match.replace(PATTERNS.ARRAYBRACKETS, "").trim();
            if (foundVals.indexOf(name.toLowerCase()) === -1) {
                foundVals.push(name.toLowerCase());
                let itmKind = vscode_1.CompletionItemKind.Variable;
                if ((/\bconst\b/i).test(matches[1]))
                    itmKind = vscode_1.CompletionItemKind.Constant;
                const ci = new vscode_1.CompletionItem(name, itmKind);
                ci.documentation = matches[3];
                if (new RegExp(PATTERNS.COLOR, "i").test(name)) {
                    ci.kind = vscode_1.CompletionItemKind.Color;
                    ci.filterText = `ColorConstants.${name}`;
                    ci.insertText = name;
                }
                ci.detail = `${matches[0]} [${scope}]`;
                CIs.push(ci);
            }
        }
    }
    return CIs;
}
function getFunctionCompletions(text, scope, parseParams = false) {
    const CIs = [];
    const foundVals = new Array();
    let matches;
    while (matches = PATTERNS.FUNCTION.exec(text)) {
        const name = matches[5];
        if (foundVals.indexOf(name.toLowerCase()) === -1) {
            foundVals.push(name.toLowerCase());
            const ci = new vscode_1.CompletionItem(name, vscode_1.CompletionItemKind.Function);
            if (matches[1]) {
                const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
                ci.documentation = summary === null || summary === void 0 ? void 0 : summary[1];
            }
            if (parseParams && matches[6])
                for (const param of matches[6].split(",")) {
                    const paramCI = new vscode_1.CompletionItem(param.trim(), vscode_1.CompletionItemKind.Variable);
                    if (matches[1]) {
                        const paramComment = PATTERNS.PARAM_SUMMARY(matches[1], param.trim());
                        if (paramComment)
                            paramCI.documentation = paramComment[1];
                    }
                    CIs.push(paramCI);
                }
            ci.detail = `${matches[2]} [${scope}]`;
            CIs.push(ci);
        }
    }
    return CIs;
}
function getPropertyCompletions(text, scope) {
    const CIs = [];
    const foundVals = new Array();
    let matches;
    while (matches = PATTERNS.PROP.exec(text)) {
        const name = matches[4];
        if (foundVals.indexOf(name.toLowerCase()) === -1) {
            foundVals.push(name.toLowerCase());
            const ci = new vscode_1.CompletionItem(name, vscode_1.CompletionItemKind.Property);
            if (matches[1]) {
                const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
                ci.documentation = summary === null || summary === void 0 ? void 0 : summary[1];
            }
            ci.detail = `${matches[2]} [${scope}]`;
            CIs.push(ci);
        }
    }
    return CIs;
}
function getClassCompletions(text, scope) {
    const CIs = [];
    const foundVals = new Array();
    let matches;
    while (matches = PATTERNS.CLASS.exec(text)) {
        const name = matches[3];
        if (foundVals.indexOf(name.toLowerCase()) === -1) {
            foundVals.push(name.toLowerCase());
            const ci = new vscode_1.CompletionItem(name, vscode_1.CompletionItemKind.Class);
            if (matches[1]) {
                const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
                ci.documentation = summary === null || summary === void 0 ? void 0 : summary[1];
            }
            ci.detail = `${name} [${scope}]`;
            CIs.push(ci);
        }
    }
    return CIs;
}
function getCompletions(text, scope, parseParams = false) {
    return [
        ...getVariableCompletions(text, scope),
        ...getFunctionCompletions(text, scope, parseParams),
        ...getPropertyCompletions(text, scope),
        ...getClassCompletions(text, scope)
    ];
}
function getObjectMembersCode(line, code, toAddObj) {
    const matches = { "Err": "Err", "WScript": "WScript", "Debug": "Debug", "fso": "FileSystemObject" };
    for (const cls of ["Err", "WScript", "Debug", "fso"]) {
        const lineClsReg = new RegExp(`.*\\b${cls}\\.\\w*`, "i");
        const codeClsReg = new RegExp(`Class[\t ]+${matches[cls]}.+?End[\t ]+Class`, "is");
        if (lineClsReg.test(line)) {
            const classDef = codeClsReg.exec(code);
            toAddObj.push(...getFunctionCompletions(classDef[0], cls), ...getPropertyCompletions(classDef[0], cls));
            return true;
        }
    }
    return false;
}
function provideCompletionItems(doc, position) {
    const codeAtPosition = doc.lineAt(position).text.substring(0, position.character);
    const line = doc.lineAt(position);
    if (line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
        return [];
    if (PATTERNS.VAR_COMPLS.test(codeAtPosition))
        return [];
    let quoteCount = 0;
    for (const char of codeAtPosition)
        if (char === '"')
            quoteCount++;
    if (quoteCount % 2 === 1)
        return [];
    const text = doc.getText();
    const retCI = [];
    const ObjectSourceImport = Includes_1.Includes.get(ObjectSourceImportName);
    const localIncludes = Includes_1.getImportsWithLocal(doc);
    if ((/.*\.\w*$/).test(codeAtPosition)) {
        if (getObjectMembersCode(codeAtPosition, ObjectSourceImport.Content, retCI)) { }
        else {
            retCI.push(...getCompletions(text, "Local"));
            retCI.push(...getFunctionCompletions(ObjectSourceImport.Content, ObjectSourceImportName), ...getPropertyCompletions(ObjectSourceImport.Content, ObjectSourceImportName));
            for (const imp of localIncludes)
                if (imp[0].startsWith("Include"))
                    retCI.push(...getCompletions(imp[1].Content, imp[0]));
        }
    }
    else {
        retCI.push(...definitions_1.default);
        retCI.push(...getCompletions(text, "Local", true));
        retCI.push(...getClassCompletions(ObjectSourceImport.Content, ObjectSourceImportName));
        for (const item of localIncludes)
            if (item[0].startsWith("Include") || item[0] === "Global")
                retCI.push(...getCompletions(item[1].Content, item[0]));
    }
    return retCI;
}
exports.default = vscode_1.languages.registerCompletionItemProvider({ scheme: "file", language: "m1s" }, { provideCompletionItems }, ".");
//# sourceMappingURL=completion.js.map