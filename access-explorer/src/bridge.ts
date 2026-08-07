import * as cp from 'node:child_process';
import * as crypto from 'node:crypto';
import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as readline from 'node:readline';
import { BridgeError } from './errors';
import { log } from './logger';

export interface DbListing {
  tables: string[];
  queries: string[];
  forms: string[];
  reports: string[];
  macros: string[];
  modules: string[];
  /** True when at least one TableDef is ODBC-linked (e.g. to SQL Server) and may need credentials. */
  hasLinkedTables: boolean;
  /** Per-object DateModified (ISO string), Modules/Forms/Reports only — lets the persistent
   *  on-disk AI mirror detect drift since it was last written without a full re-export. */
  formDates?: Record<string, string>;
  reportDates?: Record<string, string>;
  moduleDates?: Record<string, string>;
}

interface PendingRequest {
  resolve: (data: unknown) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
}

export interface BridgeOptions {
  /** Directory for request/response payload temp files. */
  workDir: string;
  /** Absolute path to ps/access-bridge.ps1. */
  scriptPath: string;
  /** Default per-operation timeout in ms. */
  defaultTimeoutMs: number;
  /** Skip the database's Startup form/AutoExec macro (classic Shift-bypass technique). */
  bypassStartup: boolean;
  /** Show Access, and the specific object being read/written, on screen in real time. */
  visibleOperations: boolean;
  /** Called when the bridge process dies unexpectedly. */
  onCrash?: (db: AccessBridge) => void;
}

const OPEN_TIMEOUT_MS = 60_000;
const COMPILE_TIMEOUT_MS = 60_000;
const COMPACT_TIMEOUT_MS = 120_000;
/**
 * Timeout for reads issued by the background AI mirror (materialize/quickSync). Mirroring is
 * fire-and-forget — nothing is waiting synchronously — so it can afford to wait longer than the
 * default interactive op timeout before a large module/report is treated as hung.
 */
export const MIRROR_TIMEOUT_MS = 45_000;
const QUIT_GRACE_MS = 5_000;
/** Margin over the poll timeout the caller requests, so the bridge watchdog never fires first. */
const UNLOCK_TIMEOUT_MARGIN_MS = 15_000;

/**
 * One long-lived PowerShell process per open database. Requests are JSON lines on
 * stdin, responses JSON lines on stdout; code bodies travel through temp files in
 * workDir so framing never depends on payload content.
 */
export class AccessBridge {
  private child: cp.ChildProcess | undefined;
  private readonly pending = new Map<number, PendingRequest>();
  private nextId = 1;
  private queue: Promise<unknown> = Promise.resolve();
  private accessPid = 0;
  private disposed = false;
  private currentOp: string | undefined;

  private constructor(
    readonly dbPath: string,
    private readonly opts: BridgeOptions,
  ) {}

  static async open(dbPath: string, opts: BridgeOptions): Promise<{ bridge: AccessBridge; vbaProtected: boolean }> {
    const bridge = new AccessBridge(dbPath, opts);
    await bridge.start();
    try {
      const data = (await bridge.request(
        'open',
        { path: dbPath, bypassStartup: opts.bypassStartup, visibleOperations: opts.visibleOperations },
        OPEN_TIMEOUT_MS,
      )) as {
        accessPid: number;
        vbaProtected?: boolean;
      };
      bridge.accessPid = data.accessPid ?? 0;
      return { bridge, vbaProtected: data.vbaProtected ?? false };
    } catch (err) {
      await bridge.dispose();
      throw err;
    }
  }

  get isAlive(): boolean {
    return !!this.child && this.child.exitCode === null && !this.disposed;
  }

  /** PID of the underlying MSACCESS.EXE this bridge is driving (0 before/if never resolved). */
  get accessProcessId(): number {
    return this.accessPid;
  }

  /** The op (and object name, if any) currently in flight — for crash diagnostics only. */
  get lastOperation(): string | undefined {
    return this.currentOp;
  }

  private async start(): Promise<void> {
    await fs.mkdir(this.opts.workDir, { recursive: true });
    this.child = cp.spawn(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-File', this.opts.scriptPath,
      ],
      { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true },
    );

    const rl = readline.createInterface({ input: this.child.stdout! });
    rl.on('line', (line) => this.onLine(line));
    this.child.stderr!.setEncoding('utf8');
    this.child.stderr!.on('data', (chunk: string) => {
      log(`[access-bridge ${path.basename(this.dbPath)}] ${chunk.trimEnd()}`);
    });
    this.child.on('exit', () => this.failAllPending(new BridgeError('BRIDGE_CRASHED')));
    this.child.on('error', (err) =>
      this.failAllPending(new BridgeError('BRIDGE_CRASHED', err.message)),
    );
  }

  private onLine(line: string): void {
    let msg: { id?: number; ok?: boolean; data?: unknown; error?: { code?: string; message?: string; number?: number } };
    try {
      msg = JSON.parse(line);
    } catch {
      log(`[access-bridge] non-JSON stdout line ignored: ${line.slice(0, 200)}`);
      return;
    }
    const entry = msg.id !== undefined ? this.pending.get(msg.id) : undefined;
    if (!entry) {
      return;
    }
    this.pending.delete(msg.id!);
    clearTimeout(entry.timer);
    if (msg.ok) {
      entry.resolve(msg.data);
    } else {
      entry.reject(new BridgeError(msg.error?.code ?? 'COM_ERROR', msg.error?.message, msg.error?.number));
    }
  }

  private failAllPending(err: Error): void {
    for (const [, entry] of this.pending) {
      clearTimeout(entry.timer);
      entry.reject(err);
    }
    this.pending.clear();
  }

  /**
   * Sends one request. Requests are serialized (Access COM is not reentrant);
   * per-op timeout triggers the hard cleanup ladder.
   */
  request(op: string, args: Record<string, unknown> = {}, timeoutMs?: number): Promise<unknown> {
    const run = async (): Promise<unknown> => {
      if (!this.isAlive) {
        throw new BridgeError('BRIDGE_CRASHED');
      }
      const id = this.nextId++;
      const timeout = timeoutMs ?? this.opts.defaultTimeoutMs;
      const nameArg = typeof args.name === 'string' ? args.name : undefined;
      this.currentOp = nameArg ? `${op}(${nameArg})` : op;
      const promise = new Promise<unknown>((resolve, reject) => {
        const timer = setTimeout(() => {
          this.pending.delete(id);
          // A hung op almost always means Access is showing a modal dialog
          // or deadlocked — the process is unusable; kill the whole ladder.
          void this.hardKill();
          reject(new BridgeError('BRIDGE_TIMEOUT', `${op} timed out after ${timeout} ms`));
        }, timeout);
        this.pending.set(id, { resolve, reject, timer });
      });
      this.child!.stdin!.write(JSON.stringify({ id, op, args }) + '\n');
      return promise;
    };
    const result = this.queue.then(run, run);
    this.queue = result.catch(() => undefined);
    return result;
  }

  // ---------- typed operations ----------

  list(): Promise<DbListing> {
    return this.request('list') as Promise<DbListing>;
  }

  /**
   * Returns the raw module export (SaveAsText, header included) or — when this
   * Access version cannot SaveAsText class modules — the VBE code body (viaVbe).
   */
  async getModule(
    name: string,
    timeoutMs?: number,
  ): Promise<{ text: string; isClass: boolean; viaVbe: boolean }> {
    return this.withTempFile(async (file) => {
      const data = (await this.request('getModule', { name, file }, timeoutMs)) as {
        isClass: boolean;
        viaVbe: boolean;
      };
      return { text: await fs.readFile(file, 'utf8'), isClass: data.isClass, viaVbe: data.viaVbe };
    });
  }

  async saveModule(name: string, fullText: string, viaVbe: boolean): Promise<void> {
    await this.withTempFile(async (file) => {
      await fs.writeFile(file, fullText, 'utf8');
      await this.request('saveModule', { name, file, viaVbe });
    });
  }

  async getMacro(name: string, timeoutMs?: number): Promise<{ text: string; enc: string }> {
    return this.withTempFile(async (file) => {
      const data = (await this.request('getMacro', { name, file }, timeoutMs)) as { enc: string };
      return { text: await fs.readFile(file, 'utf8'), enc: data.enc };
    });
  }

  async saveMacro(name: string, text: string, enc: string): Promise<void> {
    await this.withTempFile(async (file) => {
      await fs.writeFile(file, text, 'utf8');
      await this.request('saveMacro', { name, file, enc });
    });
  }

  async getQuerySql(name: string, timeoutMs?: number): Promise<string> {
    return this.withTempFile(async (file) => {
      await this.request('getQuerySql', { name, file }, timeoutMs);
      return fs.readFile(file, 'utf8');
    });
  }

  async saveQuerySql(name: string, sql: string): Promise<void> {
    await this.withTempFile(async (file) => {
      await fs.writeFile(file, sql, 'utf8');
      await this.request('saveQuerySql', { name, file });
    });
  }

  async getReadonlyDef(
    kind: 'table' | 'form' | 'report',
    name: string,
    timeoutMs?: number,
  ): Promise<{ text: string; enc?: string }> {
    const op = kind === 'table' ? 'getTableDef' : kind === 'form' ? 'getFormDef' : 'getReportDef';
    return this.withTempFile(async (file) => {
      const data = (await this.request(op, { name, file }, timeoutMs)) as { enc?: string };
      const text = await fs.readFile(file, 'utf8');
      return { text, enc: data.enc };
    });
  }

  async saveFormOrReportDef(kind: 'form' | 'report', name: string, fullText: string, enc: string): Promise<void> {
    const op = kind === 'form' ? 'saveFormDef' : 'saveReportDef';
    await this.withTempFile(async (file) => {
      await fs.writeFile(file, fullText, 'utf8');
      await this.request(op, { name, file, enc });
    });
  }

  async compile(moduleName?: string): Promise<{ compiled: boolean; message?: string }> {
    return (await this.request('compile', { name: moduleName }, COMPILE_TIMEOUT_MS)) as {
      compiled: boolean;
      message?: string;
    };
  }

  async backup(targetPath: string): Promise<void> {
    await this.request('backup', { target: targetPath }, OPEN_TIMEOUT_MS);
  }

  /**
   * Closes the database, runs Access's CompactRepair into a temp file, swaps it over the
   * original, then reopens it — all within this same bridge/process. Requires the file free
   * of any other exclusive/shared hold; throws `DB_LOCKED` (never destructive — the original
   * is left untouched and reopened) if another user/process has it open elsewhere.
   */
  async compact(): Promise<{ compacted: boolean; listing: DbListing }> {
    return (await this.request('compact', {}, COMPACT_TIMEOUT_MS)) as {
      compacted: boolean;
      listing: DbListing;
    };
  }

  /**
   * Applies the given username/password to every ODBC-linked TableDef and refreshes the link, so
   * opening a linked table's data authenticates silently instead of popping the native modal
   * login dialog (invisible to COM automation, which would hang the bridge). Throws
   * `LINKED_AUTH_FAILED` if no linked table could be relinked with these credentials.
   */
  async relinkCredentials(uid: string, pwd: string): Promise<{ relinked: number; failed: string[] }> {
    return (await this.request('relinkCredentials', { uid, pwd })) as { relinked: number; failed: string[] };
  }

  /**
   * Shows Access + the VBA editor and polls until the user has typed the project password into
   * the real Access dialog, or the timeout elapses. Never throws on timeout — `unlocked: false`
   * just means the user didn't finish in time.
   */
  async unlockVba(timeoutSeconds: number): Promise<{ unlocked: boolean }> {
    return (await this.request(
      'unlockVba',
      { timeoutSeconds },
      timeoutSeconds * 1000 + UNLOCK_TIMEOUT_MARGIN_MS,
    )) as { unlocked: boolean };
  }

  // ---------- lifecycle ----------

  /** Graceful quit, then hard kill if the bridge does not exit in time. */
  async dispose(): Promise<void> {
    if (this.disposed) {
      return;
    }
    this.disposed = true;
    const child = this.child;
    if (child && child.exitCode === null) {
      const exited = new Promise<void>((resolve) => child.once('exit', () => resolve()));
      try {
        child.stdin!.write(JSON.stringify({ id: this.nextId++, op: 'quit' }) + '\n');
      } catch {
        // stdin already closed — fall through to hard kill
      }
      const timed = await Promise.race([
        exited.then(() => true),
        new Promise<boolean>((r) => setTimeout(() => r(false), QUIT_GRACE_MS)),
      ]);
      if (!timed) {
        await this.hardKill();
      }
    }
    this.failAllPending(new BridgeError('BRIDGE_CRASHED', 'bridge disposed'));
  }

  private async hardKill(): Promise<void> {
    this.disposed = true;
    try {
      this.child?.kill();
    } catch {
      /* already dead */
    }
    if (this.accessPid > 0) {
      // Last resort: remove the orphaned MSACCESS.EXE instance.
      cp.spawn('taskkill.exe', ['/PID', String(this.accessPid), '/T', '/F'], {
        windowsHide: true,
        stdio: 'ignore',
      }).on('error', () => undefined);
      this.accessPid = 0;
    }
    // A stale lock file can survive a killed Access instance.
    const lock = this.dbPath.replace(/\.accdb$/i, '.laccdb');
    await fs.rm(lock, { force: true }).catch(() => undefined);
    this.opts.onCrash?.(this);
  }

  private async withTempFile<T>(fn: (file: string) => Promise<T>): Promise<T> {
    const file = path.join(this.opts.workDir, `payload-${crypto.randomUUID()}.txt`);
    try {
      return await fn(file);
    } finally {
      await fs.rm(file, { force: true }).catch(() => undefined);
    }
  }
}
