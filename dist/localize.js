"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Localize = void 0;
const fs_1 = require("fs");
const path_1 = require("path");
const vscode_1 = require("vscode");
class Localize {
    constructor() {
        this.bundle = this.resolveLanguagePack();
    }
    localize(key, ...args) {
        const message = this.bundle[key] || key;
        return this.format(message, args);
    }
    init() {
        this.options = {
            ...this.options,
            ...JSON.parse(process.env.VSCODE_NLS_CONFIG || "{}")
        };
    }
    format(message, args = []) {
        return args.length
            ? message.replace(/\{(\d+)\}/g, (match, rest) => args[rest[0]] || match)
            : message;
    }
    resolveLanguagePack() {
        this.init();
        const languageFormat = "package.nls{0}.json";
        const defaultLanguage = languageFormat.replace("{0}", "");
        const rootPath = vscode_1.extensions.getExtension("CalDymos.m1svscode")
            .extensionPath;
        const resolvedLanguage = this.recurseCandidates(rootPath, languageFormat, this.options.locale);
        const languageFilePath = path_1.resolve(rootPath, resolvedLanguage);
        const defaultLanguageBundle = JSON.parse(resolvedLanguage !== defaultLanguage
            ? fs_1.readFileSync(path_1.resolve(rootPath, defaultLanguage), "utf-8")
            : "{}");
        const resolvedLanguageBundle = JSON.parse(fs_1.readFileSync(languageFilePath, "utf-8"));
        return { ...defaultLanguageBundle, ...resolvedLanguageBundle };
    }
    recurseCandidates(rootPath, format, candidate) {
        const filename = format.replace("{0}", `.${candidate}`);
        const filepath = path_1.resolve(rootPath, filename);
        if (fs_1.existsSync(filepath)) {
            return filename;
        }
        if (candidate.split("-")[0] !== candidate) {
            return this.recurseCandidates(rootPath, format, candidate.split("-")[0]);
        }
        return format.replace("{0}", "");
    }
}
exports.Localize = Localize;
exports.default = Localize.prototype.localize.bind(new Localize());
//# sourceMappingURL=localize.js.map