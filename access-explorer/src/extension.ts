import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { BackupManager } from './backup';
import { AccessBridge } from './bridge';
import { describeError } from './errors';
import { AccessFsProvider } from './fsProvider';
import { Category, DatabaseRegistry, dbKeyFor } from './model';
import { AccessTreeProvider, TreeNode } from './tree';

const registry = new DatabaseRegistry();
let fsProvider: AccessFsProvider;

export function activate(context: vscode.ExtensionContext): void {
  if (process.platform !== 'win32') {
    void vscode.window.showErrorMessage(
      vscode.l10n.t('Access Explorer requires Windows with Microsoft Access installed.'),
    );
    return;
  }

  const backups = new BackupManager();
  fsProvider = new AccessFsProvider(registry, backups);
  const tree = new AccessTreeProvider(registry);

  const scriptPath = context.asAbsolutePath(path.join('ps', 'access-bridge.ps1'));
  const workDir = path.join(context.globalStorageUri.fsPath, 'work');

  context.subscriptions.push(
    fsProvider.register(),
    vscode.window.createTreeView('accessExplorer.tree', {
      treeDataProvider: tree,
      showCollapseAll: true,
    }),

    vscode.commands.registerCommand('accessExplorer.openDatabase', (uri?: vscode.Uri) =>
      openDatabase(uri, scriptPath, workDir),
    ),

    vscode.commands.registerCommand('accessExplorer.refresh', (node?: TreeNode) =>
      refresh(node && 'db' in node ? node.db.key : undefined),
    ),

    vscode.commands.registerCommand('accessExplorer.closeDatabase', async (node?: TreeNode) => {
      const key = node && 'db' in node ? node.db.key : undefined;
      if (key) {
        await closeDatabase(key);
      }
    }),

    vscode.commands.registerCommand(
      'accessExplorer.openObject',
      (key: string, category: Category, name: string) => openObject(key, category, name),
    ),

    vscode.commands.registerCommand('accessExplorer.revealBackups', async (node?: TreeNode) => {
      const db = node && 'db' in node ? node.db : undefined;
      if (db) {
        const dir = backups.backupDirFor(db);
        await fs.mkdir(dir, { recursive: true });
        void vscode.env.openExternal(vscode.Uri.file(dir));
      }
    }),
  );
}

export async function deactivate(): Promise<void> {
  // Close every bridge so no hidden MSACCESS.EXE or .laccdb outlives VS Code.
  await Promise.allSettled(registry.all.map((db) => db.bridge.dispose()));
}

async function openDatabase(
  uri: vscode.Uri | undefined,
  scriptPath: string,
  workDir: string,
): Promise<void> {
  let dbPath = uri?.fsPath;
  if (!dbPath) {
    const picked = await vscode.window.showOpenDialog({
      canSelectMany: false,
      filters: { 'Access Database': ['accdb'] },
      openLabel: vscode.l10n.t('Open in Access Explorer'),
    });
    dbPath = picked?.[0]?.fsPath;
  }
  if (!dbPath) {
    return;
  }

  const existing = registry.getByPath(dbPath);
  if (existing) {
    await refresh(existing.key);
    return;
  }

  const cfg = vscode.workspace.getConfiguration('accessExplorer');
  const timeoutMs = cfg.get<number>('operationTimeoutSeconds', 20) * 1000;

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Opening {0} in Access…', path.basename(dbPath)),
    },
    async () => {
      try {
        const bridge = await AccessBridge.open(dbPath!, {
          scriptPath,
          workDir,
          defaultTimeoutMs: timeoutMs,
          onCrash: (crashed) => onBridgeCrash(crashed.dbPath),
        });
        const listing = await bridge.list();
        registry.add({ key: dbKeyFor(dbPath!), dbPath: dbPath!, bridge, listing });
        await vscode.commands.executeCommand('accessExplorer.tree.focus');
      } catch (err) {
        void vscode.window.showErrorMessage(describeError(err));
      }
    },
  );
}

async function refresh(key?: string): Promise<void> {
  const targets = key ? [registry.get(key)].filter((d) => !!d) : registry.all;
  for (const db of targets) {
    try {
      db.listing = await db.bridge.list();
      fsProvider.invalidate(db.key);
    } catch (err) {
      void vscode.window.showErrorMessage(
        vscode.l10n.t('Refresh failed for {0}: {1}', registry.labelFor(db), describeError(err)),
      );
    }
  }
  registry.notifyChanged();
}

async function closeDatabase(key: string): Promise<void> {
  const db = registry.get(key);
  if (!db) {
    return;
  }
  registry.remove(key);
  fsProvider.dropDatabase(key);
  await db.bridge.dispose();
}

async function openObject(key: string, category: Category, name: string): Promise<void> {
  const db = registry.get(key);
  if (!db) {
    return;
  }
  try {
    const uri = await fsProvider.resolveUri(db, category, name);
    const doc = await vscode.workspace.openTextDocument(uri);
    await vscode.window.showTextDocument(doc, { preview: true });
  } catch (err) {
    void vscode.window.showErrorMessage(describeError(err));
  }
}

function onBridgeCrash(dbPath: string): void {
  const db = registry.getByPath(dbPath);
  if (db) {
    registry.remove(db.key);
    fsProvider.dropDatabase(db.key);
  }
  void vscode.window.showErrorMessage(
    vscode.l10n.t(
      'The connection to Access for {0} was lost and has been cleaned up. Reopen the database to continue.',
      path.basename(dbPath),
    ),
  );
}
