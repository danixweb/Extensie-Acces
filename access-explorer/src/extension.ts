import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { registerSelectForAi } from './aiContext';
import { AiMirrorManager } from './aiMirror';
import { BackupManager } from './backup';
import { AccessBridge } from './bridge';
import { describeError } from './errors';
import { AccessFsProvider } from './fsProvider';
import { Category, DatabaseRegistry, dbKeyFor, OpenDatabase, SCHEME } from './model';
import { AccessDbRedirectEditorProvider } from './redirectEditorProvider';
import { AccessTreeProvider, TreeNode } from './tree';
import { VbaUnlockManager } from './vbaUnlock';
import { VbaSymbolProvider } from './vbaSymbols';

const registry = new DatabaseRegistry();
let fsProvider: AccessFsProvider;
let vbaUnlock: VbaUnlockManager;
let backups: BackupManager;
let mirror: AiMirrorManager;

export function activate(context: vscode.ExtensionContext): void {
  if (process.platform !== 'win32') {
    void vscode.window.showErrorMessage(
      vscode.l10n.t('Access Explorer requires Windows with Microsoft Access installed.'),
    );
    return;
  }

  backups = new BackupManager();
  fsProvider = new AccessFsProvider(registry, backups);
  mirror = new AiMirrorManager(fsProvider, registry);
  const tree = new AccessTreeProvider(registry);
  vbaUnlock = new VbaUnlockManager(context.secrets);

  const scriptPath = context.asAbsolutePath(path.join('ps', 'access-bridge.ps1'));
  const workDir = path.join(context.globalStorageUri.fsPath, 'work');

  const treeView = vscode.window.createTreeView('accessExplorer.tree', {
    treeDataProvider: tree,
    showCollapseAll: true,
  });

  const updateFilterState = (): void => {
    void vscode.commands.executeCommand('setContext', 'accessExplorer.filterActive', tree.isFiltered);
    treeView.message = tree.isFiltered
      ? vscode.l10n.t('Filtered by: "{0}"', tree.currentFilter)
      : undefined;
  };

  context.subscriptions.push(
    fsProvider.register(),
    mirror,
    treeView,
    vscode.languages.registerDocumentSymbolProvider({ scheme: SCHEME }, new VbaSymbolProvider()),

    vscode.window.registerCustomEditorProvider(
      'accessExplorer.database',
      new AccessDbRedirectEditorProvider((uri) => openDatabase(uri, scriptPath, workDir)),
      { supportsMultipleEditorsPerDocument: false },
    ),

    vscode.commands.registerCommand('accessExplorer.openDatabase', (uri?: vscode.Uri) =>
      openDatabase(uri, scriptPath, workDir),
    ),

    vscode.commands.registerCommand('accessExplorer.refresh', (node?: TreeNode) =>
      refresh(node && 'db' in node ? node.db.key : undefined),
    ),

    vscode.commands.registerCommand('accessExplorer.remirror', (node?: TreeNode) => {
      const targets = node && 'db' in node ? [node.db] : registry.all;
      for (const db of targets) {
        materializeMirror(db);
      }
    }),

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

    vscode.commands.registerCommand('accessExplorer.revealAiMirror', async (node?: TreeNode) => {
      const db = node && 'db' in node ? node.db : undefined;
      if (db) {
        const dir = mirror.mirrorDirFor(db);
        await fs.mkdir(dir, { recursive: true });
        void vscode.env.openExternal(vscode.Uri.file(dir));
      }
    }),

    vscode.commands.registerCommand('accessExplorer.unlockVbaProject', async (node?: TreeNode) => {
      const db = node && 'db' in node ? node.db : undefined;
      if (db) {
        await vbaUnlock.unlock(db);
      }
    }),

    vscode.commands.registerCommand('accessExplorer.filterTree', async () => {
      const entered = await vscode.window.showInputBox({
        value: tree.currentFilter,
        placeHolder: vscode.l10n.t('* for everything, or text to search for'),
        prompt: vscode.l10n.t('Show only objects whose name contains this text'),
      });
      if (entered === undefined) {
        return;
      }
      tree.setFilter(entered === '*' ? '' : entered);
      updateFilterState();
    }),

    vscode.commands.registerCommand('accessExplorer.clearTreeFilter', () => {
      tree.setFilter('');
      updateFilterState();
    }),

    registerSelectForAi(),
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
  const bypassStartup = cfg.get<boolean>('bypassStartup', true);

  let openedDb: OpenDatabase | undefined;
  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Opening {0} in Access…', path.basename(dbPath)),
    },
    async () => {
      try {
        const { bridge, vbaProtected } = await AccessBridge.open(dbPath!, {
          scriptPath,
          workDir,
          defaultTimeoutMs: timeoutMs,
          bypassStartup,
          onCrash: (crashed) => onBridgeCrash(crashed.dbPath),
        });
        const listing = await bridge.list();
        const db: OpenDatabase = {
          key: dbKeyFor(dbPath!),
          dbPath: dbPath!,
          bridge,
          listing,
          vbaLocked: vbaProtected,
        };
        try {
          await backups.atSessionStart(db);
        } catch (err) {
          void vscode.window.showWarningMessage(
            vscode.l10n.t('Opened, but the session-start backup failed: {0}', describeError(err)),
          );
        }
        registry.add(db);
        openedDb = db;
        await vscode.commands.executeCommand('accessExplorer.tree.focus');
        if (vbaProtected) {
          void vbaUnlock.unlock(db);
        }
      } catch (err) {
        void vscode.window.showErrorMessage(describeError(err));
      }
    },
  );
  if (openedDb) {
    void syncMirror(openedDb);
  }
}

/** Fire-and-forget: brings the on-disk AI mirror up to date, in its own cancellable progress. Full
 *  re-export only the first time a database is mirrored; otherwise a cheap incremental check. */
function syncMirror(db: OpenDatabase): void {
  void vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Mirroring {0} objects for AI tools…', registry.labelFor(db)),
      cancellable: true,
    },
    (progress, token) => mirror.sync(db, token, progress),
  );
}

/** Fire-and-forget: forces a full re-export of every object to the on-disk AI mirror. */
function materializeMirror(db: OpenDatabase): void {
  void vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Re-mirroring {0} objects for AI tools…', registry.labelFor(db)),
      cancellable: true,
    },
    (progress, token) => mirror.materialize(db, token, progress),
  );
}

async function refresh(key?: string): Promise<void> {
  const targets = key ? [registry.get(key)].filter((d) => !!d) : registry.all;
  for (const db of targets) {
    try {
      db.listing = await db.bridge.list();
      fsProvider.invalidate(db.key);
      void syncMirror(db);
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
  await mirror.close(db);
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
    void mirror.close(db);
  }
  void vscode.window.showErrorMessage(
    vscode.l10n.t(
      'The connection to Access for {0} was lost and has been cleaned up. Reopen the database to continue.',
      path.basename(dbPath),
    ),
  );
}
