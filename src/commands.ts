import { Disposable, window, workspace } from "vscode";
import * as childProcess from "child_process";
import path from "path";
import localize from "./localize";
import * as fs from "fs";

const configuration = workspace.getConfiguration("m1s");

const m1sOut = window.createOutputChannel("Mach3 Script");

let runner: childProcess.ChildProcessWithoutNullStreams;

const scriptInterpreter: string = configuration.get<string>("interpreter");

let statbar: Disposable;

export function runScript(): void {
  if (!window.activeTextEditor)
    return;

  try {
    fs.accessSync(scriptInterpreter, fs.constants.X_OK);
  } catch {
    window.showErrorMessage(`${localize("m1s.msg.interpreterRunError")} ${ scriptInterpreter}`);
  }

  const doc = window.activeTextEditor.document;
  doc.save().then(() => {
    m1sOut.clear();
    m1sOut.show(true);

    const workDir = path.dirname(doc.fileName);

    if (statbar)
      statbar.dispose();

    statbar = window.setStatusBarMessage(localize("m1s.msg.runningscript"));

    runner = childProcess.spawn(scriptInterpreter, [doc.fileName], {
      cwd: workDir
    });

    runner.stdout.on("data", data => {
      const output = data.toString();
      m1sOut.append(output);
    });

    runner.stderr.on("data", data => {
      const output = data.toString();
      m1sOut.append(output);
    });

    runner.on("exit", code => {
      m1sOut.appendLine(`Process exited with code ${code}`);
      statbar.dispose();
    });
  }, () => {
    window.showErrorMessage("Document can' be saved");

    return;
  });
}

export function killScript(): void {
  // runner.stdin.pause();
  runner?.kill();
  statbar?.dispose();
}
