import * as vscode from 'vscode';

let channel: vscode.OutputChannel | undefined;

/** Creates the "Access Explorer" Output panel channel. Call once from activate(). */
export function initLogger(context: vscode.ExtensionContext): void {
  channel = vscode.window.createOutputChannel('Access Explorer');
  context.subscriptions.push(channel);
}

/**
 * Appends a line to the "Access Explorer" Output panel — never opens an editor tab.
 * Falls back to console.warn if called before initLogger (e.g. in tests).
 */
export function log(message: string): void {
  if (channel) {
    channel.appendLine(message);
  } else {
    console.warn(message);
  }
}
