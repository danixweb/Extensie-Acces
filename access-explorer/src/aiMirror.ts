import * as crypto from 'node:crypto';
import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { describeError } from './errors';
import { AccessFsProvider } from './fsProvider';
import {
  CATEGORIES,
  Category,
  DatabaseRegistry,
  decodeFsName,
  encodeFsName,
  extFor,
  listingFor,
  OpenDatabase,
  parseUri,
} from './model';

export const MIRROR_DIR_NAME = '.accdb-ai';

interface MirrorRecord {
  uri: vscode.Uri;
  category: Category;
  name: string;
  filePath: string;
  readonly: boolean;
  /** Last known Access-side content — what the mirror file should hold when in sync. */
  canonicalText: string;
  /** sha1 of the last content this manager itself wrote to filePath — the self-write guard. */
  lastWrittenHash: string;
}

interface DbMirrorState {
  dir: string;
  watcher: vscode.FileSystemWatcher;
  byPath: Map<string, MirrorRecord>; // keyed by filePath.toLowerCase()
  byUri: Map<string, MirrorRecord>; // keyed by uri.toString()
}

function sha1(text: string): string {
  return crypto.createHash('sha1').update(text, 'utf8').digest('hex');
}

function mirrorEnabled(): boolean {
  return vscode.workspace.getConfiguration('accessExplorer').get<boolean>('aiMirror.enabled', true);
}

/**
 * Mirrors every object of an open database to real files on disk (next to the .accdb, in
 * .accdb-ai/<Category>/<name><ext>) so external tools that only see the real filesystem — such as
 * Claude Code's @ mention and Read/Edit tools — can find and edit them. External edits are pushed
 * back into Access through AccessFsProvider.saveObject, the exact same pipeline Ctrl+S on the
 * accessdb: virtual document uses (conflict check, backup, COM write, compile check).
 */
export class AiMirrorManager implements vscode.Disposable {
  private readonly perDb = new Map<string, DbMirrorState>();
  private readonly processing = new Map<string, Promise<void>>();
  private readonly pendingRerun = new Set<string>();
  private readonly listener: vscode.Disposable;

  constructor(
    private readonly fsProvider: AccessFsProvider,
    private readonly registry: DatabaseRegistry,
  ) {
    this.listener = fsProvider.onDidChangeFile((events) => this.onVirtualDocChanged(events));
  }

  dispose(): void {
    for (const state of this.perDb.values()) {
      state.watcher.dispose();
    }
    this.perDb.clear();
    this.listener.dispose();
  }

  mirrorDirFor(db: OpenDatabase): string {
    return path.join(path.dirname(db.dbPath), MIRROR_DIR_NAME);
  }

  /** Exports the whole current listing to disk, creating the watcher on first call for this db. */
  async materialize(
    db: OpenDatabase,
    token?: vscode.CancellationToken,
    progress?: vscode.Progress<{ message?: string; increment?: number }>,
  ): Promise<void> {
    if (!mirrorEnabled()) {
      return;
    }
    const dir = this.mirrorDirFor(db);
    await fs.mkdir(dir, { recursive: true });
    const state = this.stateFor(db, dir);
    const seen = new Set<string>();
    const total = CATEGORIES.reduce((n, c) => n + listingFor(db.listing, c).length, 0) || 1;
    let done = 0;
    for (const category of CATEGORIES) {
      for (const name of listingFor(db.listing, category)) {
        if (token?.isCancellationRequested) {
          return;
        }
        try {
          await this.materializeOne(db, state, category, name, seen);
        } catch (err) {
          console.warn(`[access-explorer ai-mirror] failed to export ${category}/${name}: ${describeError(err)}`);
        }
        done++;
        progress?.report({ message: `${done} / ${total}`, increment: 100 / total });
      }
    }
    await this.pruneStale(state, seen);
  }

  /**
   * Entry point for automatic mirroring (database open, refresh): does a full `materialize()`
   * the first time a database is mirrored, otherwise a cheap `quickSync()` that never re-reads
   * unchanged objects over COM. Full re-reads only happen through the explicit re-mirror command.
   */
  async sync(
    db: OpenDatabase,
    token?: vscode.CancellationToken,
    progress?: vscode.Progress<{ message?: string; increment?: number }>,
  ): Promise<void> {
    if (!mirrorEnabled()) {
      return;
    }
    const dir = this.mirrorDirFor(db);
    if (!(await this.pathExists(dir))) {
      await this.materialize(db, token, progress);
      return;
    }
    await this.quickSync(db, dir, token, progress);
  }

  /**
   * Diffs the on-disk mirror (file names only, no COM) against the current listing to find
   * objects newly created in Access since the last mirror pass, and mirror-side files for
   * objects renamed/deleted in Access. Only genuinely new objects pay for a COM read — objects
   * already mirrored and still present are left untouched.
   */
  private async quickSync(
    db: OpenDatabase,
    dir: string,
    token?: vscode.CancellationToken,
    progress?: vscode.Progress<{ message?: string; increment?: number }>,
  ): Promise<void> {
    const state = this.stateFor(db, dir);
    const perCategory: { category: Category; added: string[]; removed: string[] }[] = [];
    for (const category of CATEGORIES) {
      const mirrored = await this.mirroredNames(dir, category);
      const current = new Set(listingFor(db.listing, category));
      const added = [...current].filter((name) => !mirrored.has(name));
      const removed = [...mirrored].filter((name) => !current.has(name));
      if (added.length > 0 || removed.length > 0) {
        perCategory.push({ category, added, removed });
      }
    }
    if (perCategory.length === 0) {
      return; // nothing new or removed — no COM reads, no disk writes
    }
    const seen = new Set<string>();
    const total = perCategory.reduce((n, c) => n + c.added.length, 0) || 1;
    let done = 0;
    for (const { category, added, removed } of perCategory) {
      for (const name of added) {
        if (token?.isCancellationRequested) {
          return;
        }
        try {
          await this.materializeOne(db, state, category, name, seen);
        } catch (err) {
          console.warn(`[access-explorer ai-mirror] failed to export ${category}/${name}: ${describeError(err)}`);
        }
        done++;
        progress?.report({ message: `${done} / ${total}`, increment: 100 / total });
      }
      for (const name of removed) {
        await this.removeMirrorFile(state, dir, category, name);
      }
    }
  }

  /** Decoded Access object names already mirrored on disk for a category — a plain directory listing, no COM. */
  private async mirroredNames(dir: string, category: Category): Promise<Set<string>> {
    const names = new Set<string>();
    let entries: string[];
    try {
      entries = await fs.readdir(path.join(dir, category));
    } catch {
      return names; // category folder doesn't exist yet — nothing mirrored
    }
    for (const entry of entries) {
      const decoded = this.decodeMirrorFileName(category, entry);
      if (decoded) {
        names.add(decoded);
      }
    }
    return names;
  }

  /** Reverses materializeOne's `encodeFsName(name) + ext` back into the original Access object name. */
  private decodeMirrorFileName(category: Category, fileName: string): string | undefined {
    const ext = extFor(category, fileName);
    if (!ext) {
      return undefined; // e.g. a stray ".prev" snapshot file
    }
    return decodeFsName(fileName.slice(0, fileName.length - ext.length));
  }

  /** Reverses a mirror file's on-disk path back into (category, name), for watcher events that arrive as bare paths. */
  private parseMirrorFilePath(state: DbMirrorState, filePath: string): { category: Category; name: string } | undefined {
    const segments = path.relative(state.dir, filePath).split(path.sep);
    if (segments.length !== 2 || !CATEGORIES.includes(segments[0] as Category)) {
      return undefined;
    }
    const category = segments[0] as Category;
    const name = this.decodeMirrorFileName(category, segments[1]);
    return name ? { category, name } : undefined;
  }

  private async removeMirrorFile(state: DbMirrorState, dir: string, category: Category, name: string): Promise<void> {
    // The extension isn't known without a COM read, but readdir gave us the exact file name we scanned.
    let entries: string[];
    try {
      entries = await fs.readdir(path.join(dir, category));
    } catch {
      return;
    }
    for (const entry of entries) {
      if (this.decodeMirrorFileName(category, entry) !== name) {
        continue;
      }
      const filePath = path.join(dir, category, entry);
      const key = filePath.toLowerCase();
      const record = state.byPath.get(key);
      if (record) {
        state.byPath.delete(key);
        state.byUri.delete(record.uri.toString());
      }
      await fs.rm(filePath, { force: true }).catch(() => undefined);
      await fs.rm(filePath + '.prev', { force: true }).catch(() => undefined);
    }
  }

  private async pathExists(p: string): Promise<boolean> {
    try {
      await fs.access(p);
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Lazily populates a MirrorRecord for an object that quickSync left untouched (already
   * mirrored, unchanged), via a single COM read — used the first time such an object is edited
   * from either side after a quick, incremental open. Never overwrites the on-disk mirror file,
   * so a pending external edit is preserved for the caller to push back to Access.
   */
  private async ensureRecord(
    db: OpenDatabase,
    state: DbMirrorState,
    category: Category,
    name: string,
  ): Promise<MirrorRecord | undefined> {
    const existingByName = [...state.byPath.values()].find((r) => r.category === category && r.name === name);
    if (existingByName) {
      return existingByName;
    }
    try {
      const { uri, text, readonly } = await this.fsProvider.readObject(db, category, name);
      const fileName = uri.path.slice(uri.path.lastIndexOf('/') + 1);
      const ext = extFor(category, fileName);
      if (!ext) {
        return undefined;
      }
      const resolvedPath = path.join(state.dir, category, encodeFsName(name) + ext);
      const key = resolvedPath.toLowerCase();
      const record: MirrorRecord = {
        uri,
        category,
        name,
        filePath: resolvedPath,
        readonly,
        canonicalText: text,
        lastWrittenHash: sha1(text),
      };
      state.byPath.set(key, record);
      state.byUri.set(uri.toString(), record);
      return record;
    } catch (err) {
      console.warn(`[access-explorer ai-mirror] failed to hydrate ${category}/${name}: ${describeError(err)}`);
      return undefined;
    }
  }

  /** Stops watching and removes the mirror folder — it is a derived artifact, not a durable store. */
  async close(db: OpenDatabase): Promise<void> {
    const state = this.perDb.get(db.key);
    if (!state) {
      return;
    }
    this.perDb.delete(db.key);
    state.watcher.dispose();
    await fs.rm(state.dir, { recursive: true, force: true }).catch(() => undefined);
  }

  // ---------- materialize ----------

  private stateFor(db: OpenDatabase, dir: string): DbMirrorState {
    const existing = this.perDb.get(db.key);
    if (existing) {
      return existing;
    }
    const dbKey = db.key;
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.file(dir), '**/*'));
    watcher.onDidChange((uri) => this.onDiskEvent(dbKey, uri));
    watcher.onDidCreate((uri) => this.onDiskEvent(dbKey, uri));
    watcher.onDidDelete((uri) => this.onDiskDeleteEvent(dbKey, uri));
    const state: DbMirrorState = { dir, watcher, byPath: new Map(), byUri: new Map() };
    this.perDb.set(dbKey, state);
    return state;
  }

  private async materializeOne(
    db: OpenDatabase,
    state: DbMirrorState,
    category: Category,
    name: string,
    seen: Set<string>,
  ): Promise<void> {
    const { uri, text, readonly } = await this.fsProvider.readObject(db, category, name);
    const fileName = uri.path.slice(uri.path.lastIndexOf('/') + 1);
    const ext = extFor(category, fileName);
    if (!ext) {
      return; // defensive — resolveUri always produces a matching extension
    }
    const filePath = path.join(state.dir, category, encodeFsName(name) + ext);
    const key = filePath.toLowerCase();
    seen.add(key);

    const hash = sha1(text);
    const existing = state.byPath.get(key);
    if (existing && existing.lastWrittenHash === hash) {
      return; // unchanged since last materialize — skip disk IO
    }
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await this.writeMirrorFile(filePath, text);
    const record: MirrorRecord = { uri, category, name, filePath, readonly, canonicalText: text, lastWrittenHash: hash };
    state.byPath.set(key, record);
    state.byUri.set(uri.toString(), record);
  }

  /** Removes mirror files for objects renamed/deleted in Access since the last materialize. */
  private async pruneStale(state: DbMirrorState, seen: Set<string>): Promise<void> {
    for (const [key, record] of [...state.byPath]) {
      if (seen.has(key)) {
        continue;
      }
      state.byPath.delete(key);
      state.byUri.delete(record.uri.toString());
      await fs.rm(record.filePath, { force: true }).catch(() => undefined);
      await fs.rm(record.filePath + '.prev', { force: true }).catch(() => undefined);
    }
  }

  private async writeMirrorFile(filePath: string, text: string): Promise<void> {
    await fs.writeFile(filePath, text, 'utf8');
  }

  // ---------- virtual doc (accessdb:) -> mirror, keeps both views in sync ----------

  private onVirtualDocChanged(events: vscode.FileChangeEvent[]): void {
    for (const event of events) {
      if (event.type !== vscode.FileChangeType.Changed) {
        continue;
      }
      void this.syncFromVirtualDoc(event.uri).catch((err) =>
        console.warn(`[access-explorer ai-mirror] mirror sync failed: ${describeError(err)}`),
      );
    }
  }

  private async syncFromVirtualDoc(uri: vscode.Uri): Promise<void> {
    let parsed;
    try {
      parsed = parseUri(uri);
    } catch {
      return;
    }
    const state = this.perDb.get(parsed.key);
    const db = this.registry.get(parsed.key);
    if (!state || !db) {
      return;
    }
    const record = state.byUri.get(uri.toString()) ?? (await this.ensureRecord(db, state, parsed.category, parsed.name));
    if (!record) {
      return;
    }
    const fresh = await this.fsProvider.readObject(db, record.category, record.name);
    if (fresh.text === record.canonicalText) {
      return;
    }
    await this.writeMirrorFile(record.filePath, fresh.text);
    record.canonicalText = fresh.text;
    record.lastWrittenHash = sha1(fresh.text);
  }

  // ---------- disk watcher: external edit -> save-to-Access -> revert ----------

  /** Serializes handling per path and coalesces a second event that arrives mid-processing. */
  private schedule(key: string, run: () => Promise<void>): void {
    if (this.processing.has(key)) {
      this.pendingRerun.add(key);
      return;
    }
    const runOnce = (async (): Promise<void> => {
      try {
        await run();
      } catch (err) {
        console.warn(`[access-explorer ai-mirror] ${key}: ${describeError(err)}`);
      } finally {
        this.processing.delete(key);
        if (this.pendingRerun.delete(key)) {
          this.schedule(key, run);
        }
      }
    })();
    this.processing.set(key, runOnce);
  }

  private onDiskEvent(dbKey: string, uri: vscode.Uri): void {
    const filePath = uri.fsPath;
    if (filePath.toLowerCase().endsWith('.prev')) {
      return; // our own revert-snapshot files, never treated as object edits
    }
    this.schedule(filePath.toLowerCase(), () => this.handleDiskChange(dbKey, filePath));
  }

  private onDiskDeleteEvent(dbKey: string, uri: vscode.Uri): void {
    const filePath = uri.fsPath;
    if (filePath.toLowerCase().endsWith('.prev')) {
      return;
    }
    this.schedule(filePath.toLowerCase(), () => this.handleDiskDelete(dbKey, filePath));
  }

  private async handleDiskChange(dbKey: string, filePath: string): Promise<void> {
    if (await this.isDirectory(filePath)) {
      return; // category folders (e.g. "Modules") fire their own watcher events; nothing to sync
    }
    const state = this.perDb.get(dbKey);
    const db = this.registry.get(dbKey);
    let record = state?.byPath.get(filePath.toLowerCase());
    if (!record && state && db) {
      // Not yet hydrated — this may be an object quickSync left untouched (already mirrored,
      // unchanged) rather than an unknown file. Confirm it against the current listing first.
      const parsed = this.parseMirrorFilePath(state, filePath);
      if (parsed && listingFor(db.listing, parsed.category).includes(parsed.name)) {
        record = await this.ensureRecord(db, state, parsed.category, parsed.name);
      }
    }
    if (!state || !record) {
      void vscode.window.showWarningMessage(
        vscode.l10n.t('"{0}" in the AI mirror does not match a known database object; ignoring.', path.basename(filePath)),
      );
      return;
    }
    if (!db) {
      return; // database was closed — close() already tore this mirror down
    }
    let content: string;
    try {
      content = await fs.readFile(filePath, 'utf8');
    } catch {
      return; // file briefly missing mid-write (e.g. atomic replace); a following event retries
    }
    const hash = sha1(content);
    if (hash === record.lastWrittenHash) {
      return; // identical to what we last wrote ourselves — not an external edit
    }
    if (record.readonly) {
      await this.writeMirrorFile(record.filePath, record.canonicalText);
      record.lastWrittenHash = sha1(record.canonicalText);
      void vscode.window.showWarningMessage(
        vscode.l10n.t(
          '"{0}" ({1}) is read-only in the AI mirror — your change was discarded. Edit it in Access instead.',
          record.name,
          record.category,
        ),
      );
      return;
    }

    const prevFilePath = `${record.filePath}.prev`;
    try {
      await fs.writeFile(prevFilePath, record.canonicalText, 'utf8');
      const fresh = await this.fsProvider.saveObject(record.uri, content);
      await this.writeMirrorFile(record.filePath, fresh.text);
      record.canonicalText = fresh.text;
      record.lastWrittenHash = sha1(fresh.text);
      const revert = vscode.l10n.t('Revert');
      const choice = await vscode.window.showInformationMessage(
        vscode.l10n.t('AI edit saved to Access: {0} ({1}). Test it live in Access.', record.name, record.category),
        revert,
      );
      if (choice === revert) {
        await this.revert(db, record);
      }
    } catch (err) {
      // Nothing was actually applied — resync the mirror to the true current content and drop the snapshot.
      this.fsProvider.invalidateOne(record.uri);
      try {
        const fresh = await this.fsProvider.readObject(db, record.category, record.name);
        await this.writeMirrorFile(record.filePath, fresh.text);
        record.canonicalText = fresh.text;
        record.lastWrittenHash = sha1(fresh.text);
      } catch {
        /* best-effort resync only */
      }
      await fs.rm(prevFilePath, { force: true }).catch(() => undefined);
      void vscode.window.showErrorMessage(describeError(err, { vbaLocked: db.vbaLocked }));
    }
  }

  private async isDirectory(filePath: string): Promise<boolean> {
    try {
      return (await fs.stat(filePath)).isDirectory();
    } catch {
      return false; // gone or inaccessible — let the caller's normal handling decide
    }
  }

  private async handleDiskDelete(dbKey: string, filePath: string): Promise<void> {
    const state = this.perDb.get(dbKey);
    const db = this.registry.get(dbKey);
    let record = state?.byPath.get(filePath.toLowerCase());
    if (!record && state && db) {
      const parsed = this.parseMirrorFilePath(state, filePath);
      if (parsed && listingFor(db.listing, parsed.category).includes(parsed.name)) {
        record = await this.ensureRecord(db, state, parsed.category, parsed.name);
      }
    }
    if (!record) {
      return;
    }
    // There is no bridge operation to delete an Access object — restore the file instead of
    // silently losing the mirror's tracking of it.
    await this.writeMirrorFile(record.filePath, record.canonicalText);
    record.lastWrittenHash = sha1(record.canonicalText);
    void vscode.window.showWarningMessage(
      vscode.l10n.t(
        'Access Explorer does not delete database objects through the AI mirror; "{0}" has been restored.',
        record.name,
      ),
    );
  }

  private async revert(db: OpenDatabase, record: MirrorRecord): Promise<void> {
    const prevFilePath = `${record.filePath}.prev`;
    let prevText: string;
    try {
      prevText = await fs.readFile(prevFilePath, 'utf8');
    } catch {
      return; // nothing to revert to
    }
    try {
      const fresh = await this.fsProvider.saveObject(record.uri, prevText);
      await this.writeMirrorFile(record.filePath, fresh.text);
      record.canonicalText = fresh.text;
      record.lastWrittenHash = sha1(fresh.text);
      await fs.rm(prevFilePath, { force: true }).catch(() => undefined);
      void vscode.window.showInformationMessage(vscode.l10n.t('Reverted {0} ({1}).', record.name, record.category));
    } catch (err) {
      void vscode.window.showErrorMessage(describeError(err, { vbaLocked: db.vbaLocked }));
    }
  }
}
