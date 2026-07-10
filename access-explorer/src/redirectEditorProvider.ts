import * as vscode from 'vscode';

export type OpenDatabaseFn = (uri: vscode.Uri) => Promise<void>;

/**
 * Claims .accdb as the "default" custom editor so VS Code never shows its built-in binary
 * editor for it (double-click from Windows, drag-and-drop, `code file.accdb`). It never really
 * renders a webview: it hands off to Access Explorer's own openDatabase/tree flow and disposes
 * its own panel on the next tick, since there's no API to claim a file type without VS Code
 * momentarily creating a webview panel for it.
 */
export class AccessDbRedirectEditorProvider implements vscode.CustomReadonlyEditorProvider {
  constructor(private readonly openDatabase: OpenDatabaseFn) {}

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose(): void {} };
  }

  resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel): void {
    webviewPanel.webview.options = { enableScripts: false };
    webviewPanel.webview.html = loadingHtml();

    void this.openDatabase(document.uri);

    setTimeout(() => {
      try {
        webviewPanel.dispose();
      } catch {
        // Already disposed (e.g. user closed the tab manually mid-flight) — nothing to do.
      }
    }, 0);
  }
}

function loadingHtml(): string {
  return `<!DOCTYPE html><html><body style="font-family:sans-serif;padding:1em;">${vscode.l10n.t('Opening in Access Explorer…')}</body></html>`;
}
