import * as cp from 'node:child_process';
import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { registerSelectForAi } from './aiContext';
import { AiMirrorManager } from './aiMirror';
import { BackupManager } from './backup';
import { AccessBridge } from './bridge';
import { describeError } from './errors';
import { AccessFsProvider } from './fsProvider';
import { parseDesignText } from './formDesignParser';
import { generateHtml } from './htmlMockupGenerator';
import { LinkedCredentialsManager } from './linkedCredentials';
import { initLogger, log } from './logger';
import { Category, DatabaseRegistry, dbKeyFor, encodeFsName, OpenDatabase, parseUri, SCHEME } from './model';
import { AccessDbRedirectEditorProvider } from './redirectEditorProvider';
import { AccessTreeProvider, TreeNode } from './tree';
import { VbaUnlockManager } from './vbaUnlock';
import { VbaSymbolProvider } from './vbaSymbols';

const registry = new DatabaseRegistry();
let fsProvider: AccessFsProvider;
let vbaUnlock: VbaUnlockManager;
let linkedCredentials: LinkedCredentialsManager;
let backups: BackupManager;
let mirror: AiMirrorManager;

export function activate(context: vscode.ExtensionContext): void {
  if (process.platform !== 'win32') {
    void vscode.window.showErrorMessage(
      vscode.l10n.t('Access Explorer requires Windows with Microsoft Access installed.'),
    );
    return;
  }

  initLogger(context);
  backups = new BackupManager();
  fsProvider = new AccessFsProvider(registry, backups);
  mirror = new AiMirrorManager(fsProvider, registry);
  const tree = new AccessTreeProvider(registry);
  vbaUnlock = new VbaUnlockManager(context.secrets);
  linkedCredentials = new LinkedCredentialsManager(context.secrets);

  const scriptPath = context.asAbsolutePath(path.join('ps', 'access-bridge.ps1'));
  const findOrphansScriptPath = context.asAbsolutePath(path.join('ps', 'find-orphaned-access.ps1'));
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
      new AccessDbRedirectEditorProvider((uri) => openDatabase(uri, scriptPath, workDir, findOrphansScriptPath)),
      { supportsMultipleEditorsPerDocument: false },
    ),

    vscode.commands.registerCommand('accessExplorer.openDatabase', (uri?: vscode.Uri) =>
      openDatabase(uri, scriptPath, workDir, findOrphansScriptPath),
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

    vscode.commands.registerCommand('accessExplorer.compactDatabase', async (node?: TreeNode) => {
      const db = node && 'db' in node ? node.db : await pickOpenDatabase();
      if (db) {
        await compactDatabase(db, true);
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

    vscode.commands.registerCommand('accessExplorer.exportHtmlMockup', async (node?: TreeNode) => {
      const target = resolveFormOrReportTarget(node);
      if (!target) {
        void vscode.window.showErrorMessage(
          vscode.l10n.t('Open a Form or Report first, or run this from its item in the tree.'),
        );
        return;
      }
      await exportHtmlMockup(target.db, target.category, target.name);
    }),
  );
}

/** Resolves the Form/Report to export either from a tree selection or the active accessdb: editor. */
function resolveFormOrReportTarget(
  node?: TreeNode,
): { db: OpenDatabase; category: Category; name: string } | undefined {
  if (node && node.kind === 'object') {
    return { db: node.db, category: node.category, name: node.name };
  }
  const editor = vscode.window.activeTextEditor;
  if (!editor || editor.document.uri.scheme !== SCHEME) {
    return undefined;
  }
  try {
    const { key, category, name } = parseUri(editor.document.uri);
    const db = registry.get(key);
    return db ? { db, category, name } : undefined;
  } catch {
    return undefined;
  }
}

/** Generates a self-contained HTML mockup of a Form/Report's design and opens it in the browser. */
async function exportHtmlMockup(db: OpenDatabase, category: Category, name: string): Promise<void> {
  if (category !== 'Forms' && category !== 'Reports') {
    void vscode.window.showErrorMessage(
      vscode.l10n.t('HTML mockup export only works on Forms and Reports.'),
    );
    return;
  }
  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Generating HTML mockup for {0}…', name),
    },
    async () => {
      try {
        const raw = await fsProvider.getDesignSection(db, category, name);
        const parsed = parseDesignText(raw);
        if (!parsed) {
          void vscode.window.showErrorMessage(
            vscode.l10n.t(
              'Could not recognize the design section for {0} — its export format may differ from what this command expects.',
              name,
            ),
          );
          return;
        }
        const html = generateHtml(parsed, name);
        const dir = path.join(mirror.mirrorDirFor(db), category);
        await fs.mkdir(dir, { recursive: true });
        const outPath = path.join(dir, encodeFsName(name) + '.mockup.html');
        await fs.writeFile(outPath, html, 'utf8');
        void vscode.env.openExternal(vscode.Uri.file(outPath));
        void vscode.window.showInformationMessage(vscode.l10n.t('HTML mockup written to {0}', outPath));
      } catch (err) {
        void vscode.window.showErrorMessage(describeError(err));
      }
    },
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
  findOrphansScriptPath: string,
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

  await checkOrphanedAccessProcesses(findOrphansScriptPath);

  const cfg = vscode.workspace.getConfiguration('accessExplorer');
  const timeoutMs = cfg.get<number>('operationTimeoutSeconds', 20) * 1000;
  const bypassStartup = cfg.get<boolean>('bypassStartup', true);
  const visibleOperations = cfg.get<boolean>('visibleOperations', false);

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
          visibleOperations,
          onCrash: (crashed) => onBridgeCrash(crashed.dbPath, crashed.lastOperation),
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
    await linkedCredentials.ensureConnected(openedDb);
    await mirror.ensureDir(openedDb);
  }
}

interface OrphanCandidate {
  pid: number;
  commandLine: string;
  startTime: string | null;
}

function execFile(file: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    cp.execFile(file, args, { windowsHide: true }, (err, stdout) => (err ? reject(err) : resolve(stdout)));
  });
}

/**
 * Best-effort check for MSACCESS.EXE processes left over from a crashed/killed automation session
 * (this extension's own bridge, or a standalone dev script) — never blocks opening a database over
 * this. Databases this extension currently has open legitimately are excluded by PID, so only truly
 * unaccounted-for processes are ever offered for closing, and only at the user's explicit choice.
 */
async function checkOrphanedAccessProcesses(findOrphansScriptPath: string): Promise<void> {
  let candidates: OrphanCandidate[];
  try {
    const stdout = await execFile('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', findOrphansScriptPath,
    ]);
    candidates = JSON.parse(stdout) as OrphanCandidate[];
  } catch {
    return;
  }
  const excludePids = new Set(registry.all.map((db) => db.bridge.accessProcessId).filter((pid) => pid > 0));
  const toReview = candidates.filter((c) => !excludePids.has(c.pid));
  if (toReview.length === 0) {
    return;
  }

  const items = toReview.map((c) => ({
    label: vscode.l10n.t('PID {0}', c.pid),
    description: c.startTime ?? undefined,
    detail: c.commandLine,
    pid: c.pid,
  }));
  const picked = await vscode.window.showQuickPick(items, {
    canPickMany: true,
    placeHolder: vscode.l10n.t(
      'Found {0} orphaned Access process(es) (no window, likely left over from a crashed session). Select any to close, or press Escape to leave them.',
      toReview.length,
    ),
  });
  if (!picked || picked.length === 0) {
    return;
  }

  const results = await Promise.allSettled(
    picked.map((item) => execFile('taskkill.exe', ['/PID', String(item.pid), '/T', '/F'])),
  );
  const closed = results.filter((r) => r.status === 'fulfilled').length;
  const failed = results.length - closed;
  if (failed === 0) {
    void vscode.window.showInformationMessage(vscode.l10n.t('Closed {0} orphaned Access process(es).', closed));
  } else {
    void vscode.window.showWarningMessage(
      vscode.l10n.t('Closed {0} orphaned Access process(es); {1} failed to close.', closed, failed),
    );
  }
}

/** Fire-and-forget: re-reads only the objects already present in the on-disk AI mirror (e.g. after
 *  "Refresh") — objects the user never opened are never pulled in; see AiMirrorManager.mirrorOnDemand
 *  for how new objects get mirrored (on tree click, plus their true dependencies). */
function resyncMirror(db: OpenDatabase): void {
  void vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Refreshing mirrored objects for {0}…', registry.labelFor(db)),
      cancellable: true,
    },
    (progress, token) => mirror.resyncMirrored(db, token, progress),
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
      resyncMirror(db);
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
  const cfg = vscode.workspace.getConfiguration('accessExplorer');
  if (cfg.get<boolean>('compactOnClose', true)) {
    // Best-effort: a failure here (e.g. another user/process has the file open elsewhere,
    // so CompactRepair can't get the access it needs) never blocks closing.
    await compactDatabase(db, false);
  }
  registry.remove(key);
  fsProvider.dropDatabase(key);
  await mirror.close(db);
  await db.bridge.dispose();
}

async function pickOpenDatabase(): Promise<OpenDatabase | undefined> {
  if (registry.all.length === 0) {
    void vscode.window.showInformationMessage(vscode.l10n.t('No database is currently open in Access Explorer.'));
    return undefined;
  }
  if (registry.all.length === 1) {
    return registry.all[0];
  }
  const picked = await vscode.window.showQuickPick(
    registry.all.map((db) => ({ label: registry.labelFor(db), db })),
    { placeHolder: vscode.l10n.t('Select the database to compact') },
  );
  return picked?.db;
}

/** Closes the database, runs CompactRepair, and reopens it — all within the same bridge/process.
 *  Never throws: a failure (e.g. locked by another user) surfaces as a warning and leaves the
 *  database open and untouched. */
async function compactDatabase(db: OpenDatabase, notifyOnSuccess: boolean): Promise<boolean> {
  let ok = false;
  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: vscode.l10n.t('Compacting {0}…', registry.labelFor(db)),
    },
    async () => {
      try {
        const { listing } = await db.bridge.compact();
        db.listing = listing;
        fsProvider.invalidate(db.key);
        registry.notifyChanged();
        ok = true;
        if (notifyOnSuccess) {
          void vscode.window.showInformationMessage(vscode.l10n.t('Compacted {0}.', registry.labelFor(db)));
        }
      } catch (err) {
        void vscode.window.showWarningMessage(
          vscode.l10n.t('Compact failed for {0}: {1}', registry.labelFor(db), describeError(err)),
        );
      }
    },
  );
  return ok;
}

async function openObject(key: string, category: Category, name: string): Promise<void> {
  const db = registry.get(key);
  if (!db) {
    return;
  }
  void mirror.mirrorOnDemand(db, category, name);
  try {
    const uri = await fsProvider.resolveUri(db, category, name);
    const doc = await vscode.workspace.openTextDocument(uri);
    await vscode.window.showTextDocument(doc, { preview: true });
  } catch (err) {
    void vscode.window.showErrorMessage(describeError(err));
  }
}

function onBridgeCrash(dbPath: string, lastOperation?: string): void {
  const db = registry.getByPath(dbPath);
  if (db) {
    registry.remove(db.key);
    fsProvider.dropDatabase(db.key);
    void mirror.close(db);
  }
  log(
    `[access-bridge] connection to ${path.basename(dbPath)} was lost` +
      (lastOperation ? ` while running: ${lastOperation}` : ''),
  );
  void vscode.window.showErrorMessage(
    vscode.l10n.t(
      'The connection to Access for {0} was lost and has been cleaned up. Reopen the database to continue.',
      path.basename(dbPath),
    ),
  );
}
