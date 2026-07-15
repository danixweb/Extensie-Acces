import * as vscode from 'vscode';
import { AiMirrorManager } from './aiMirror';
import { Category, DatabaseRegistry, SCHEME, parseUri } from './model';
import { CODELESS_CATEGORIES, groupProcedures } from './vbaSymbols';

/**
 * Sets the selection to exactly the code the user wants an AI assistant to focus on.
 * The accessdb: virtual document (where the object is normally viewed/edited) is invisible to
 * external tools such as Claude Code, which only see real files on disk — so the selection has to
 * land in the object's AI-mirror file (see aiMirror.ts) instead, since that's the file such tools
 * actually read.
 */
export function registerSelectForAi(registry: DatabaseRegistry, mirror: AiMirrorManager): vscode.Disposable {
  return vscode.commands.registerCommand('accessExplorer.selectForAi', () => selectForAi(registry, mirror));
}

interface PickItem extends vscode.QuickPickItem {
  range?: vscode.Range;
}

async function selectForAi(registry: DatabaseRegistry, mirror: AiMirrorManager): Promise<void> {
  const editor = vscode.window.activeTextEditor;
  if (!editor || editor.document.uri.scheme !== SCHEME) {
    void vscode.window.showInformationMessage(
      vscode.l10n.t('Open an Access object first — this command only works on Access Explorer documents.'),
    );
    return;
  }
  const document = editor.document;

  let parsed;
  try {
    parsed = parseUri(document.uri);
  } catch {
    return;
  }
  const { key, category, name } = parsed;

  const wholeDocument = fullDocumentRange(document);
  let range = wholeDocument;

  if (!CODELESS_CATEGORIES.has(category)) {
    const groups = groupProcedures(document);
    if (groups.length > 0) {
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
      if (!picked) {
        return;
      }
      range = picked.range ?? wholeDocument;
    }
  }

  await focusInMirror(registry, mirror, key, category, name, range, editor);
}

/** Opens (or reuses) the real AI-mirror file for this object and selects `range` there, so the
 *  selection is visible to external tools. Falls back to selecting in the virtual document —
 *  still useful as visual feedback in-editor — if the mirror is unavailable (e.g. disabled, or
 *  the database connection is gone). */
async function focusInMirror(
  registry: DatabaseRegistry,
  mirror: AiMirrorManager,
  key: string,
  category: Category,
  name: string,
  range: vscode.Range,
  fallbackEditor: vscode.TextEditor,
): Promise<void> {
  const db = registry.get(key);
  let filePath: string | undefined;
  if (db) {
    try {
      filePath = await mirror.ensureMirrorFile(db, category, name);
    } catch {
      filePath = undefined;
    }
  }

  if (!filePath) {
    applySelection(fallbackEditor, range);
    void vscode.window.showWarningMessage(
      vscode.l10n.t(
        'Could not open the AI mirror file for "{0}" — selected in the Access Explorer editor instead, which external AI tools cannot see.',
        name,
      ),
    );
    return;
  }

  const mirrorDocument = await vscode.workspace.openTextDocument(vscode.Uri.file(filePath));
  const mirrorEditor = await vscode.window.showTextDocument(mirrorDocument, { preview: false });
  applySelection(mirrorEditor, range);
}

function fullDocumentRange(document: vscode.TextDocument): vscode.Range {
  const lastLine = document.lineCount - 1;
  return new vscode.Range(0, 0, lastLine, document.lineAt(lastLine).text.length);
}

function applySelection(editor: vscode.TextEditor, range: vscode.Range): void {
  editor.selection = new vscode.Selection(range.start, range.end);
  editor.revealRange(range, vscode.TextEditorRevealType.InCenter);
}
