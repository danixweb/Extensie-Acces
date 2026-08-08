import * as vscode from 'vscode';
import { DatabaseRegistry, OpenDatabase, parseUri } from './model';
import { parseProcedures } from './vbaSymbols';
import { MIRROR_DIR_NAME } from './aiMirror';

export class PushToAccessCodeLensProvider implements vscode.CodeLensProvider {
  constructor(registry: DatabaseRegistry) {}

  provideCodeLenses(
    document: vscode.TextDocument,
    token: vscode.CancellationToken
  ): vscode.ProviderResult<vscode.CodeLens[]> {
    const isMirror = document.uri.scheme === 'file' && document.uri.fsPath.includes(MIRROR_DIR_NAME);
    const isVirtual = document.uri.scheme === 'accessdb';
    if (!isMirror && !isVirtual) {
      return [];
    }

    const procs = parseProcedures(document.getText());
    const lenses: vscode.CodeLens[] = [];

    for (const proc of procs) {
      const range = new vscode.Range(proc.startLine, 0, proc.endLine, 0);
      const command: vscode.Command = {
        title: 'Transfera VBE',
        command: 'accessExplorer.pushChunk',
        arguments: [document.uri, proc.name, proc.kindWord, range]
      };
      lenses.push(new vscode.CodeLens(range, command));
    }

    return lenses;
  }
}

export function registerPushToAccess(context: vscode.ExtensionContext, registry: DatabaseRegistry) {
  const provider = new PushToAccessCodeLensProvider(registry);

  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      [{ scheme: 'accessdb' }, { scheme: 'file', language: 'vba' }, { scheme: 'file', language: 'plaintext' }],
      provider
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand(
      'accessExplorer.pushChunk',
      async (uri: vscode.Uri, procName: string, kindWord: string, range: vscode.Range) => {
        let db: OpenDatabase | undefined;
        let moduleName = '';

        if (uri.scheme === 'accessdb') {
          const parsed = parseUri(uri);
          if (!parsed) return;
          db = registry.get(parsed.key);
          moduleName = parsed.name;
          if (parsed.category === 'Forms') moduleName = `Form_${moduleName}`;
          if (parsed.category === 'Reports') moduleName = `Report_${moduleName}`;
        } else {
          // If mirror, we pick the first open DB for simplicity, as we usually have 1.
          const dbs = registry.all;
          if (dbs.length === 0) {
            vscode.window.showErrorMessage('No active database connection found.');
            return;
          }
          db = dbs[0];

          const match = uri.fsPath.match(/[\\/](Forms|Reports|Modules)[\\/](.+?)(?:\.form\.txt|\.report\.txt|\.cls|\.bas)$/i);
          if (match) {
            const category = match[1].toLowerCase();
            const name = match[2];
            if (category === 'forms') moduleName = `Form_${name}`;
            else if (category === 'reports') moduleName = `Report_${name}`;
            else moduleName = name;
          } else {
            vscode.window.showErrorMessage('Could not determine Access module name from file path.');
            return;
          }
        }

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.toString() !== uri.toString()) {
          vscode.window.showErrorMessage('File must be active in editor to push code.');
          return;
        }

        const text = editor.document.getText(range);

        try {
          await vscode.window.withProgress(
            { location: vscode.ProgressLocation.Notification, title: `Transfera VBE: ${procName}...` },
            async () => {
              await db!.bridge.replaceVbeChunk(moduleName, procName, kindWord, text);
            }
          );
          vscode.window.showInformationMessage(`Transferat cu succes în VBE: ${procName}`);
        } catch (e: any) {
          vscode.window.showErrorMessage(`Eroare VBE: ${e.message}`);
        }
      }
    )
  );
}
