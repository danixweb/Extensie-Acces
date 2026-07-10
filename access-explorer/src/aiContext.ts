import * as vscode from 'vscode';
import { SCHEME, parseUri } from './model';
import { CODELESS_CATEGORIES, groupProcedures } from './vbaSymbols';

/**
 * Sets the editor selection to exactly the code the user wants an AI assistant to focus on.
 * Selected text (unlike the document's virtual `accessdb:` URI) reaches chat tools directly,
 * so this is how "look at this query" / "look at this control's event" actually works.
 */
export function registerSelectForAi(): vscode.Disposable {
  return vscode.commands.registerCommand('accessExplorer.selectForAi', selectForAi);
}

interface PickItem extends vscode.QuickPickItem {
  range?: vscode.Range;
}

async function selectForAi(): Promise<void> {
  const editor = vscode.window.activeTextEditor;
  if (!editor || editor.document.uri.scheme !== SCHEME) {
    void vscode.window.showInformationMessage(
      vscode.l10n.t('Open an Access object first — this command only works on Access Explorer documents.'),
    );
    return;
  }
  const document = editor.document;

  let category;
  try {
    ({ category } = parseUri(document.uri));
  } catch {
    return;
  }

  const wholeDocument = fullDocumentRange(document);

  if (CODELESS_CATEGORIES.has(category)) {
    // Tables/Queries/Macros have no procedures to narrow down to — select everything.
    applySelection(editor, wholeDocument);
    return;
  }

  const groups = groupProcedures(document);
  if (groups.length === 0) {
    applySelection(editor, wholeDocument);
    return;
  }

  const items: PickItem[] = [
    { label: vscode.l10n.t('$(book) Whole module code'), range: wholeDocument },
  ];
  for (const group of groups) {
    items.push({ label: group.bucket, kind: vscode.QuickPickItemKind.Separator });
    items.push({
      label: vscode.l10n.t('$(symbol-class) All events of {0}', group.bucket),
      range: group.range,
    });
    for (const event of group.events) {
      items.push({ label: `    $(arrow-small-right) ${event.label}`, range: event.range });
    }
  }

  const picked = await vscode.window.showQuickPick(items, {
    placeHolder: vscode.l10n.t('Select the code to focus the AI on'),
  });
  if (picked?.range) {
    applySelection(editor, picked.range);
  }
}

function fullDocumentRange(document: vscode.TextDocument): vscode.Range {
  const lastLine = document.lineCount - 1;
  return new vscode.Range(0, 0, lastLine, document.lineAt(lastLine).text.length);
}

function applySelection(editor: vscode.TextEditor, range: vscode.Range): void {
  editor.selection = new vscode.Selection(range.start, range.end);
  editor.revealRange(range, vscode.TextEditorRevealType.InCenter);
}
