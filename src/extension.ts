import { commands, ExtensionContext, workspace } from "vscode";
import hoverProvider from "./hover";
import completionProvider from "./completion";
import symbolsProvider from "./symbols";
import signatureProvider from "./signature";
import definitionProvider from "./definition";
import colorProvider from "./colorprovider";
import launchProvider from "./Launcher";
import * as cmds from "./commands";
import { IncludeFile, Includes, reloadImportDocuments } from "./Includes";

export function activate(context: ExtensionContext): void {
  Includes.set("Global", new IncludeFile(context.asAbsolutePath("./GlobalDefs.m1s")));
  Includes.set("ObjectDefs", new IncludeFile(context.asAbsolutePath("./ObjectDefs.m1s")));

  workspace.onDidChangeConfiguration(reloadImportDocuments);
  reloadImportDocuments();

  context.subscriptions.push(
    hoverProvider,
    completionProvider,
    symbolsProvider,
    signatureProvider,
    definitionProvider,
    colorProvider,
    launchProvider.launchConfigProvider,
    launchProvider.inlineDebugAdapterFactory
  );

  // Run Script Command
  commands.registerCommand("m1s.runScript", () => {
    cmds.runScript();
  });

  // Kill running script command
  commands.registerCommand("m1s.killScript", () => {
    cmds.killScript();
  });
}
