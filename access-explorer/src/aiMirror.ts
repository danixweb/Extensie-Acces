import * as crypto from 'node:crypto';
import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { DbListing, MIRROR_TIMEOUT_MS } from './bridge';
import { describeError } from './errors';
import { AccessFsProvider } from './fsProvider';
import { log } from './logger';
import {
  extractCalledProcedureNames,
  extractDomainReferences,
  extractSqlReferences,
} from './mirrorDependencies';
import {
  CATEGORIES,
  Category,
  DatabaseRegistry,
  decodeFsName,
  encodeFsName,
  extFor,
  listingFor,
  objectUri,
  OpenDatabase,
  parseUri,
} from './model';
import { parseProcedures } from './vbaSymbols';

export const MIRROR_DIR_NAME = '.accdb-ai';

/** Sidecar recording, per object, the Access-side DateModified last seen when it was mirrored —
 *  lets a persisted mirror file from a previous session be trusted without a COM re-fetch. */
const MIRROR_META_FILE = '.mirror-meta.json';
/** Manifest of the last on-demand mirror pass: the clicked object plus its discovered
 *  dependencies, so an external tool can jump straight to the relevant files. */
const WORKING_SET_FILE = '.working-set.json';
/** Written by cursorContext.ts: which routine the user's cursor is currently inside, in whichever
 *  accessdb: document is active — lets an external tool infer the intended target of a command
 *  like /vba-analiza without an explicit name or a real (mirror-file) selection. */
export const CURSOR_CONTEXT_FILE = '.cursor-context.json';

/** Categories with a cheap DateModified signal (see access-bridge.ps1 Op-List) — the only ones
 *  primeFromDisk can validate without a full re-export; Tables/Queries/Macros always refetch. */
const DATE_TRACKED_CATEGORIES: ReadonlySet<Category> = new Set(['Modules', 'Forms', 'Reports']);

/** "<Category>/<UPPER-NAME>" key used by the .mirror-meta.json sidecar (Access names are
 *  case-insensitive). */
function metaKey(category: Category, name: string): string {
  return `${category}/${name.toUpperCase()}`;
}

function dateMapFor(listing: DbListing, category: Category): Record<string, string> | undefined {
  switch (category) {
    case 'Modules':
      return listing.moduleDates;
    case 'Forms':
      return listing.formDates;
    case 'Reports':
      return listing.reportDates;
    default:
      return undefined;
  }
}

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
  /** metaKey(category, name) -> last known Access-side DateModified, persisted to MIRROR_META_FILE. */
  meta: Record<string, string>;
}

function sha1(text: string): string {
  return crypto.createHash('sha1').update(text, 'utf8').digest('hex');
}

function mirrorEnabled(): boolean {
  return vscode.workspace.getConfiguration('accessExplorer').get<boolean>('aiMirror.enabled', true);
}

/** Pause between objects during background mirroring, so a tight loop of COM reads doesn't
 *  pile up unreleased references faster than the bridge process can reclaim them. */
const MIRROR_STEP_DELAY_MS = 75;

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
  /** db.key set once primeFromDisk has run for it this VS Code session. */
  private readonly primed = new Set<string>();
  private readonly listener: vscode.Disposable;
  /** Per-db procedure-name (uppercased) -> module-name index, built lazily by ensureModuleIndex. */
  private readonly procIndex = new Map<string, Map<string, string>>();
  private readonly procIndexBuilding = new Map<string, Promise<Map<string, string>>>();

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
    await this.ensureStatePrimed(db, state);
    const seen = new Set<string>();
    const total = CATEGORIES.reduce((n, c) => n + listingFor(db.listing, c).length, 0) || 1;
    let done = 0;
    for (const category of CATEGORIES) {
      for (const name of listingFor(db.listing, category)) {
        if (token?.isCancellationRequested) {
          return;
        }
        if (!db.bridge.isAlive) {
          log(`[ai-mirror] mirroring aborted — connection lost (${done}/${total} done)`);
          return;
        }
        try {
          await this.materializeOne(db, state, category, name, seen, { forceRefresh: true });
          log(`[ai-mirror] ${category}/${name}: OK (${done + 1}/${total})`);
        } catch (err) {
          log(`[ai-mirror] ${category}/${name}: FAILED — ${describeError(err)} (${done + 1}/${total})`);
        }
        done++;
        progress?.report({ message: `${done} / ${total}`, increment: 100 / total });
        await delay(MIRROR_STEP_DELAY_MS);
      }
    }
    await this.pruneStale(state, seen);
  }

  /** Ensures the mirror dir + watcher exist for a just-opened database, with no COM reads —
   *  objects are mirrored on demand (see `mirrorOnDemand`), not eagerly at open. Also primes
   *  byPath/byUri from any mirror files persisted from a previous session (primeFromDisk), which
   *  is itself COM-free — it only compares against db.listing's already-fetched DateModified. */
  async ensureDir(db: OpenDatabase): Promise<void> {
    if (!mirrorEnabled()) {
      return;
    }
    const dir = this.mirrorDirFor(db);
    await fs.mkdir(dir, { recursive: true });
    const state = this.stateFor(db, dir);
    await this.ensureStatePrimed(db, state);
  }

  /**
   * Refreshes only the objects already mirrored (e.g. after "Refresh" re-read the listing) —
   * never pulls in objects the user hasn't opened, unlike the old eager full-listing sync.
   */
  async resyncMirrored(
    db: OpenDatabase,
    token?: vscode.CancellationToken,
    progress?: vscode.Progress<{ message?: string; increment?: number }>,
  ): Promise<void> {
    if (!mirrorEnabled()) {
      return;
    }
    const dir = this.mirrorDirFor(db);
    let state = this.perDb.get(db.key);
    if (!state) {
      // First touch this session: sets up watching and primes byPath/byUri from any mirror files
      // persisted from a previous session, so a "Refresh" right after reopening a database still
      // finds (and force-refreshes) whatever was already mirrored, not just this-session state.
      await this.ensureDir(db);
      state = this.perDb.get(db.key);
    }
    if (!state) {
      return; // mirroring disabled
    }
    await this.ensureStatePrimed(db, state);
    const records = [...state.byPath.values()];
    const seen = new Set<string>();
    const total = records.length || 1;
    let done = 0;
    for (const record of records) {
      if (token?.isCancellationRequested) {
        return;
      }
      if (!db.bridge.isAlive) {
        log(`[ai-mirror] refresh aborted — connection lost (${done}/${total} done)`);
        return;
      }
      if (!listingFor(db.listing, record.category).some((n) => n.toUpperCase() === record.name.toUpperCase())) {
        await this.removeMirrorFile(state, dir, record.category, record.name);
        done++;
        continue;
      }
      try {
        await this.materializeOne(db, state, record.category, record.name, seen, { forceRefresh: true });
        log(`[ai-mirror] ${record.category}/${record.name}: OK (${done + 1}/${total})`);
      } catch (err) {
        log(`[ai-mirror] ${record.category}/${record.name}: FAILED — ${describeError(err)} (${done + 1}/${total})`);
      }
      done++;
      progress?.report({ message: `${done} / ${total}`, increment: 100 / total });
      await delay(MIRROR_STEP_DELAY_MS);
    }
  }

  /**
   * Mirrors one object (typically: the one just clicked in the Access Explorer tree) plus its
   * true dependencies, recursively: tables/queries referenced by a query's SQL, tables/queries
   * named as a D-function domain, and other modules containing procedures the code calls. This is
   * the on-demand replacement for eagerly mirroring the whole database at open — a database with
   * 1770+ objects would otherwise time out/degrade long before finishing (see MIRROR_TIMEOUT_MS
   * and the COM-cleanup work in access-bridge.ps1 for that earlier, now-secondary, class of fix).
   */
  async mirrorOnDemand(db: OpenDatabase, category: Category, name: string): Promise<void> {
    if (!mirrorEnabled()) {
      return;
    }
    const dir = this.mirrorDirFor(db);
    await fs.mkdir(dir, { recursive: true });
    const state = this.stateFor(db, dir);
    await this.ensureStatePrimed(db, state);
    const seen = new Set<string>(); // satisfies materializeOne's signature; no pruneStale pass here
    const visited = new Set<string>();
    const queue: { category: Category; name: string }[] = [{ category, name }];
    const workingSet: { category: Category; name: string; file: string; role: 'focus' | 'dependency' }[] = [];

    while (queue.length > 0) {
      const next = queue.shift()!;
      const visitKey = `${next.category} ${next.name.toUpperCase()}`;
      if (visited.has(visitKey)) {
        continue;
      }
      visited.add(visitKey);
      const resolvedName = listingFor(db.listing, next.category).find(
        (n) => n.toUpperCase() === next.name.toUpperCase(),
      );
      if (!resolvedName) {
        continue; // renamed/deleted since the listing was last refreshed
      }
      if (!db.bridge.isAlive) {
        log(`[ai-mirror] on-demand mirroring aborted — connection lost`);
        return;
      }

      let text: string;
      try {
        text = await this.materializeOne(db, state, next.category, resolvedName, seen);
        log(`[ai-mirror] ${next.category}/${resolvedName}: OK (on-demand)`);
        const record = [...state.byPath.values()].find(
          (r) => r.category === next.category && r.name === resolvedName,
        );
        if (record) {
          workingSet.push({
            category: next.category,
            name: resolvedName,
            file: path.relative(dir, record.filePath).split(path.sep).join('/'),
            role: workingSet.length === 0 ? 'focus' : 'dependency',
          });
        }
      } catch (err) {
        log(`[ai-mirror] ${next.category}/${resolvedName}: FAILED — ${describeError(err)} (on-demand)`);
        continue;
      }
      await delay(MIRROR_STEP_DELAY_MS);

      const dependencyNames = new Set<string>();
      if (next.category === 'Queries') {
        for (const ref of extractSqlReferences(text)) {
          dependencyNames.add(ref);
        }
      } else if (next.category === 'Modules' || next.category === 'Forms' || next.category === 'Reports') {
        for (const ref of extractDomainReferences(text)) {
          dependencyNames.add(ref);
        }
        const localProcNames = new Set(parseProcedures(text).map((p) => p.name.toUpperCase()));
        const hasExternalCandidate = [...text.matchAll(/\b[A-Za-z_][A-Za-z0-9_]*\s*\(/g)].some(
          (m) => !localProcNames.has(m[0].replace(/\s*\($/, '').toUpperCase()),
        );
        if (hasExternalCandidate) {
          const index = await this.ensureModuleIndex(db, state);
          for (const upper of extractCalledProcedureNames(text, new Set(index.keys()))) {
            const moduleName = index.get(upper);
            if (moduleName) {
              queue.push({ category: 'Modules', name: moduleName });
            }
          }
        }
      }
      for (const depName of dependencyNames) {
        const resolved = this.resolveTableOrQuery(db, depName);
        if (resolved) {
          queue.push(resolved);
        }
      }
    }
    await this.writeWorkingSetManifest(dir, workingSet);
  }

  /**
   * Ensures a single object (no dependency expansion) is mirrored to disk and returns its real
   * file path. Used by "Select Code for AI": the accessdb: virtual document editor is invisible
   * to external tools like Claude Code, which only see real files, so making a selection stick
   * requires opening/selecting in this mirror file rather than the virtual one.
   */
  async ensureMirrorFile(db: OpenDatabase, category: Category, name: string): Promise<string | undefined> {
    if (!mirrorEnabled()) {
      return undefined;
    }
    const dir = this.mirrorDirFor(db);
    await fs.mkdir(dir, { recursive: true });
    const state = this.stateFor(db, dir);
    await this.ensureStatePrimed(db, state);
    const seen = new Set<string>(); // satisfies materializeOne's signature; no pruneStale pass here
    await this.materializeOne(db, state, category, name, seen);
    const record = [...state.byPath.values()].find((r) => r.category === category && r.name === name);
    return record?.filePath;
  }

  /** Overwrites WORKING_SET_FILE with the focus object + dependencies just discovered by
   *  mirrorOnDemand's BFS, so an external tool (e.g. Claude Code) can jump straight to the
   *  relevant mirrored files instead of scanning the whole .accdb-ai/ tree. Reflects only the
   *  most recent click — always overwritten, never appended. */
  private async writeWorkingSetManifest(
    dir: string,
    objects: { category: Category; name: string; file: string; role: 'focus' | 'dependency' }[],
  ): Promise<void> {
    const payload = { generatedAt: new Date().toISOString(), objects };
    await fs
      .writeFile(path.join(dir, WORKING_SET_FILE), JSON.stringify(payload, null, 2), 'utf8')
      .catch(() => undefined);
  }

  /** Matches a name (from SQL/domain-function extraction) against Tables first, then Queries — both
   *  are valid D-function/SQL domains, case-insensitively, since Access names are case-insensitive. */
  private resolveTableOrQuery(db: OpenDatabase, name: string): { category: Category; name: string } | undefined {
    for (const category of ['Tables', 'Queries'] as const) {
      const match = listingFor(db.listing, category).find((n) => n.toUpperCase() === name.toUpperCase());
      if (match) {
        return { category, name: match };
      }
    }
    return undefined;
  }

  /**
   * Lazily builds (once per open database) a procedure-name → module-name index across every
   * Module/Form/Report, so `mirrorOnDemand` can find which *other* module a called procedure lives
   * in. This is the one remaining "read everything in a category" operation in the new on-demand
   * design — but deferred until the first module with an unresolved call is actually opened,
   * rather than running unconditionally at every database open, and it reuses the same safety nets
   * (isAlive abort, MIRROR_TIMEOUT_MS, pacing delay) as the old eager mirror. Each module read here
   * is also written to the on-disk mirror via materializeOne, so the cost is paid once.
   */
  private async ensureModuleIndex(db: OpenDatabase, state: DbMirrorState): Promise<Map<string, string>> {
    const existing = this.procIndex.get(db.key);
    if (existing) {
      return existing;
    }
    const building = this.procIndexBuilding.get(db.key);
    if (building) {
      return building;
    }
    const promise = this.buildModuleIndex(db, state);
    this.procIndexBuilding.set(db.key, promise);
    try {
      const index = await promise;
      this.procIndex.set(db.key, index);
      return index;
    } finally {
      this.procIndexBuilding.delete(db.key);
    }
  }

  private async buildModuleIndex(db: OpenDatabase, state: DbMirrorState): Promise<Map<string, string>> {
    const index = new Map<string, string>();
    const moduleNames = listingFor(db.listing, 'Modules');
    const seen = new Set<string>();
    const total = moduleNames.length || 1;
    let done = 0;
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: vscode.l10n.t('Indexing modules for {0}…', path.basename(db.dbPath)),
        cancellable: true,
      },
      async (progress, token) => {
        for (const name of moduleNames) {
          if (token.isCancellationRequested || !db.bridge.isAlive) {
            log(`[ai-mirror] module index build stopped (${done}/${total} done)`);
            return;
          }
          try {
            const text = await this.materializeOne(db, state, 'Modules', name, seen);
            for (const proc of parseProcedures(text)) {
              index.set(proc.name.toUpperCase(), name);
            }
          } catch (err) {
            log(`[ai-mirror] failed to index Modules/${name}: ${describeError(err)}`);
          }
          done++;
          progress.report({ message: `${done} / ${total}`, increment: 100 / total });
          await delay(MIRROR_STEP_DELAY_MS);
        }
      },
    );
    return index;
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

  /**
   * Lazily populates a MirrorRecord for an object not mirrored yet (either never opened, or from
   * an on-demand pass that stopped before reaching it), via a single COM read — used the first
   * time such an object is edited from either side. Never overwrites the on-disk mirror file,
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
      log(`[ai-mirror] failed to hydrate ${category}/${name}: ${describeError(err)}`);
      return undefined;
    }
  }

  /**
   * Stops watching and waits for any sync already in flight to finish — the mirror directory
   * itself is kept on disk as a cross-session cache (see primeFromDisk), no longer wiped here.
   * Must run, and be awaited, before the caller releases the COM bridge (registry.remove +
   * bridge.dispose in extension.ts's closeDatabase), otherwise an in-flight push-back would find
   * the database already gone and fail instead of completing.
   */
  async close(db: OpenDatabase): Promise<void> {
    this.procIndex.delete(db.key);
    this.procIndexBuilding.delete(db.key);
    this.primed.delete(db.key);
    const state = this.perDb.get(db.key);
    if (!state) {
      return;
    }
    state.watcher.dispose();
    await this.flushPending(state.dir);
    this.perDb.delete(db.key);
  }

  /**
   * Awaits any handleDiskChange/handleDiskDelete runs already in flight for paths under `dir`.
   * processing/pendingRerun are keyed by bare file path, not by database, so filter by prefix.
   * Loops because a coalesced rerun (pendingRerun) can start a fresh promise for the same key
   * right as the one we're awaiting resolves.
   */
  private async flushPending(dir: string): Promise<void> {
    const prefix = dir.toLowerCase();
    for (;;) {
      const matching = [...this.processing.entries()].filter(([key]) => key.startsWith(prefix));
      if (matching.length === 0) {
        return;
      }
      await Promise.all(matching.map(([, p]) => p));
    }
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
    const state: DbMirrorState = { dir, watcher, byPath: new Map(), byUri: new Map(), meta: {} };
    this.perDb.set(dbKey, state);
    return state;
  }

  private async loadMirrorMeta(dir: string): Promise<Record<string, string>> {
    try {
      const raw = await fs.readFile(path.join(dir, MIRROR_META_FILE), 'utf8');
      return JSON.parse(raw) as Record<string, string>;
    } catch {
      return {};
    }
  }

  private async persistMirrorMeta(state: DbMirrorState): Promise<void> {
    await fs
      .writeFile(path.join(state.dir, MIRROR_META_FILE), JSON.stringify(state.meta, null, 2), 'utf8')
      .catch(() => undefined);
  }

  private recordMeta(state: DbMirrorState, category: Category, name: string, dateModified: string | undefined): void {
    if (!dateModified || !DATE_TRACKED_CATEGORIES.has(category)) {
      return;
    }
    state.meta[metaKey(category, name)] = dateModified;
  }

  /** Runs primeFromDisk exactly once per database per VS Code session, the first time its mirror
   *  state is touched (materialize/mirrorOnDemand/ensureDir/resyncMirrored all call this right
   *  after stateFor). */
  private async ensureStatePrimed(db: OpenDatabase, state: DbMirrorState): Promise<void> {
    if (this.primed.has(db.key)) {
      return;
    }
    this.primed.add(db.key);
    await this.primeFromDisk(db, state);
  }

  /**
   * Populates byPath/byUri from mirror files already on disk from a previous session — with zero
   * COM calls — for objects whose recorded DateModified (MIRROR_META_FILE) still matches the
   * fresh value from this session's db.listing (see access-bridge.ps1 Op-List). Objects with no
   * recorded date, or a drifted one, are left unprimed: materializeOne will treat them exactly
   * like a never-mirrored object and do a real COM fetch on first touch, overwriting the stale
   * file. Scoped to Modules/Forms/Reports — the only categories with a DateModified signal.
   */
  private async primeFromDisk(db: OpenDatabase, state: DbMirrorState): Promise<void> {
    state.meta = await this.loadMirrorMeta(state.dir);
    for (const category of DATE_TRACKED_CATEGORIES) {
      let entries: string[];
      try {
        entries = await fs.readdir(path.join(state.dir, category));
      } catch {
        continue; // nothing mirrored in this category yet
      }
      const freshDates = dateMapFor(db.listing, category);
      for (const entry of entries) {
        if (entry.endsWith('.prev')) {
          continue;
        }
        const name = this.decodeMirrorFileName(category, entry);
        if (!name) {
          continue;
        }
        const resolvedName = listingFor(db.listing, category).find((n) => n.toUpperCase() === name.toUpperCase());
        if (!resolvedName) {
          continue; // renamed/deleted since last session — a future Refresh will clean it up
        }
        const knownDate = state.meta[metaKey(category, resolvedName)];
        const currentDate = freshDates?.[resolvedName];
        if (!knownDate || !currentDate || knownDate !== currentDate) {
          continue; // unknown or drifted — treat as not-yet-mirrored
        }
        const filePath = path.join(state.dir, category, entry);
        let content: string;
        try {
          content = await fs.readFile(filePath, 'utf8');
        } catch {
          continue;
        }
        const ext = extFor(category, entry);
        if (!ext) {
          continue;
        }
        const uri = objectUri(db.key, category, resolvedName, ext);
        const record: MirrorRecord = {
          uri,
          category,
          name: resolvedName,
          filePath,
          readonly: false,
          canonicalText: content,
          lastWrittenHash: sha1(content),
        };
        state.byPath.set(filePath.toLowerCase(), record);
        state.byUri.set(uri.toString(), record);
      }
    }
  }

  /** Returns the object's fetched text so callers that also need to parse it (on-demand dependency
   *  expansion, module index build) don't pay for a second COM round-trip. When an object is
   *  already known (this session, or primed from a previous one) and `forceRefresh` isn't set,
   *  trusts it as-is with zero COM cost — the "mirror only what isn't mirrored yet" path.
   *  `forceRefresh` (used by the explicit Refresh/"re-mirror everything" commands) always
   *  re-fetches from Access, still skipping the disk write if nothing actually changed. */
  private async materializeOne(
    db: OpenDatabase,
    state: DbMirrorState,
    category: Category,
    name: string,
    seen: Set<string>,
    opts: { forceRefresh?: boolean } = {},
  ): Promise<string> {
    const existingByName = [...state.byPath.values()].find((r) => r.category === category && r.name === name);
    if (existingByName && !opts.forceRefresh) {
      seen.add(existingByName.filePath.toLowerCase());
      return existingByName.canonicalText;
    }

    const { uri, text, readonly } = await this.fsProvider.readObject(db, category, name, MIRROR_TIMEOUT_MS);
    const fileName = uri.path.slice(uri.path.lastIndexOf('/') + 1);
    const ext = extFor(category, fileName);
    if (!ext) {
      return text; // defensive — resolveUri always produces a matching extension
    }
    const filePath = path.join(state.dir, category, encodeFsName(name) + ext);
    const key = filePath.toLowerCase();
    seen.add(key);

    const hash = sha1(text);
    if (existingByName && existingByName.lastWrittenHash === hash) {
      return text; // forced refresh confirmed nothing changed — skip disk IO
    }
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await this.writeMirrorFile(filePath, text);
    const record: MirrorRecord = { uri, category, name, filePath, readonly, canonicalText: text, lastWrittenHash: hash };
    state.byPath.set(key, record);
    state.byUri.set(uri.toString(), record);
    this.recordMeta(state, category, name, dateMapFor(db.listing, category)?.[name]);
    void this.persistMirrorMeta(state);
    return text;
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
        log(`[ai-mirror] mirror sync failed: ${describeError(err)}`),
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
        log(`[ai-mirror] ${key}: ${describeError(err)}`);
      } finally {
        this.processing.delete(key);
        if (this.pendingRerun.delete(key)) {
          this.schedule(key, run);
        }
      }
    })();
    this.processing.set(key, runOnce);
  }

  /** True for our own sidecar files (.prev revert snapshots, the meta/working-set manifests) —
   *  never treated as database object edits by the watcher. */
  private isSidecarPath(filePath: string): boolean {
    if (filePath.toLowerCase().endsWith('.prev')) {
      return true;
    }
    const base = path.basename(filePath);
    return base === MIRROR_META_FILE || base === WORKING_SET_FILE || base === CURSOR_CONTEXT_FILE;
  }

  private onDiskEvent(dbKey: string, uri: vscode.Uri): void {
    const filePath = uri.fsPath;
    if (this.isSidecarPath(filePath)) {
      return;
    }
    this.schedule(filePath.toLowerCase(), () => this.handleDiskChange(dbKey, filePath));
  }

  private onDiskDeleteEvent(dbKey: string, uri: vscode.Uri): void {
    const filePath = uri.fsPath;
    if (this.isSidecarPath(filePath)) {
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
      // Not yet hydrated — this may be a real, existing database object that on-demand mirroring
      // just hasn't reached yet, rather than an unknown file. Confirm it against the listing first.
      const parsed = this.parseMirrorFilePath(state, filePath);
      if (parsed) {
        const resolvedName = listingFor(db.listing, parsed.category).find(
          (n) => n.toUpperCase() === parsed.name.toUpperCase()
        );
        if (resolvedName) {
          record = await this.ensureRecord(db, state, parsed.category, resolvedName);
        }
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
      record.canonicalText = fresh.text;

      // A further edit may have landed on disk while the COM round-trip above was in
      // flight; schedule() will already have queued a rerun for it (pendingRerun). If we
      // blindly overwrote the mirror file with this save's canonical form now, we'd clobber
      // that newer edit before it's ever read, and the rerun's hash check would then treat
      // it as our own echo and silently drop it. Only normalize the file when nothing else
      // has arrived since we read `content`.
      let diskNow: string;
      try {
        diskNow = await fs.readFile(filePath, 'utf8');
      } catch {
        diskNow = content;
      }
      if (sha1(diskNow) === hash) {
        await this.writeMirrorFile(record.filePath, fresh.text);
        record.lastWrittenHash = sha1(fresh.text);
      }
      // Access stamped a fresh DateModified via the DoCmd.Save inside saveObject — record "now"
      // as an approximation so a future session's primeFromDisk knows this file is caught up.
      this.recordMeta(state, record.category, record.name, new Date().toISOString());
      void this.persistMirrorMeta(state);

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
      if (parsed) {
        const resolvedName = listingFor(db.listing, parsed.category).find(
          (n) => n.toUpperCase() === parsed.name.toUpperCase()
        );
        if (resolvedName) {
          record = await this.ensureRecord(db, state, parsed.category, resolvedName);
        }
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
