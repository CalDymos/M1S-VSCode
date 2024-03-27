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
const vscode_1 = require("vscode");
const Includes_1 = require("./Includes");
const PATTERNS = __importStar(require("./patterns"));
function getParsableCode(code) {
    const reducedCode = code
        .replace(/\w+\([^()]*\)/g, "")
        .replace(/"[^"]*"/g, "")
        .replace(/"[^"]*(?=$)/g, "")
        .replace(/\([^()]*\)/g, "")
        .replace(/\({2,}/g, "(");
    return reducedCode;
}
function getCurrentFunction(code) {
    const parenSplit = code.split("(");
    let index;
    if (parenSplit.length === 1)
        index = 0;
    else
        index = parenSplit.length - 2;
    return parenSplit[index].match(/(?:.*)\b(\w+)/)[1].toLowerCase();
}
function countCommas(code) {
    const openParen = code.lastIndexOf("(");
    const commas = code.slice(openParen).match(/(?!\B["'][^"']*),(?![^"']*['"]\B)/g);
    if (commas === null) {
        return 0;
    }
    else {
        return commas.length;
    }
}
function getCallInfo(doc, pos) {
    const codeAtPosition = doc.lineAt(pos).text.substring(0, pos.character);
    const cleanCode = getParsableCode(codeAtPosition);
    return {
        func: getCurrentFunction(cleanCode),
        commas: countCommas(cleanCode)
    };
}
function getSignatures(text, docComment) {
    var _a;
    const map = new Map();
    let matches;
    while ((matches = PATTERNS.FUNCTION.exec(text)) !== null) {
        const name = matches[5].toLowerCase();
        let documentation = "";
        if (matches[1]) {
            const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
            if (summary)
                documentation = summary[1];
            else
                documentation = docComment;
        }
        const si = new vscode_1.SignatureInformation(matches[4], documentation);
        (_a = matches[6]) === null || _a === void 0 ? void 0 : _a.split(",").forEach(param => {
            const pi = new vscode_1.ParameterInformation(param.trim());
            if (matches[1]) {
                const paramComment = PATTERNS.PARAM_SUMMARY(matches[1], param.trim());
                if (paramComment)
                    pi.documentation = paramComment[1];
            }
            si.parameters.push(pi);
        });
        let prevMatches;
        if ((prevMatches = map.get(name)) !== undefined)
            map.set(name, [
                ...prevMatches,
                si
            ]);
        else
            map.set(name, [si]);
    }
    return map;
}
function provideSignatureHelp(doc, position, _token, context) {
    const caller = getCallInfo(doc, position);
    if (caller === null)
        return null;
    const sighelp = new vscode_1.SignatureHelp();
    if (context.activeSignatureHelp)
        sighelp.activeSignature = context.activeSignatureHelp.activeSignature;
    else
        sighelp.activeSignature = 0;
    sighelp.activeParameter = caller.commas;
    let sig;
    if ((sig = getSignatures(doc.getText(), "Local").get(caller.func)) !== undefined) {
        sighelp.signatures.push(...sig.filter((sig2) => sig2.parameters.length >= caller.commas));
    }
    for (const item of Includes_1.getImportsWithLocal(doc)) {
        if ((sig = getSignatures(item[1].Content, item[0]).get(caller.func)) !== undefined) {
            sighelp.signatures.push(...sig.filter((sig2) => sig2.parameters.length >= caller.commas));
        }
    }
    return sighelp;
}
exports.default = vscode_1.languages.registerSignatureHelpProvider({ language: "m1s" }, { provideSignatureHelp }, "(", ",", " ");
//# sourceMappingURL=signature.js.map