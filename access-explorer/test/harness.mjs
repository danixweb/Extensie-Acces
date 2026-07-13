// harness.mjs — headless client for ps/access-bridge.ps1, used by run-tests.mjs.
//
// Mirrors the production client (src/bridge.ts): same spawn arguments, JSON-lines
// protocol, one in-flight request, temp-file payloads, and the same kill ladder
// (kill powershell -> taskkill MSACCESS pid -> delete .laccdb) so a hung Access
// (hidden modal dialog) can never wedge a test run.

import { spawn, execFile } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as readline from 'node:readline';

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export async function waitFor(cond, timeoutMs, intervalMs = 250) {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    if (await cond()) return true;
    if (Date.now() >= deadline) return false;
    await sleep(intervalMs);
  }
}

export function pidAlive(pid) {
  if (!pid) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err.code === 'EPERM';
  }
}

export function taskkill(pid, { tree = false } = {}) {
  return new Promise((resolve) => {
    const args = ['/PID', String(pid), ...(tree ? ['/T'] : []), '/F'];
    execFile('taskkill', args, { windowsHide: true }, () => resolve());
  });
}

/** PIDs of all running processes with the given image name (e.g. 'MSACCESS.EXE'). */
export function listProcesses(imageName) {
  return new Promise((resolve) => {
    execFile(
      'tasklist',
      ['/FI', `IMAGENAME eq ${imageName}`, '/FO', 'CSV', '/NH'],
      { windowsHide: true },
      (err, stdout) => {
        if (err || !stdout) return resolve([]);
        const pids = [];
        for (const line of stdout.split(/\r?\n/)) {
          const m = line.match(/^"([^"]+)","(\d+)"/);
          if (m && m[1].toLowerCase() === imageName.toLowerCase()) pids.push(Number(m[2]));
        }
        resolve(pids);
      },
    );
  });
}

export function laccdbPath(dbPath) {
  const ext = path.extname(dbPath);
  const lockExt = ext.toLowerCase() === '.mdb' ? '.ldb' : '.laccdb';
  return dbPath.slice(0, dbPath.length - ext.length) + lockExt;
}

export async function removeLaccdb(dbPath, attempts = 3) {
  const lock = laccdbPath(dbPath);
  for (let i = 0; i < attempts; i++) {
    if (!fs.existsSync(lock)) return true;
    try {
      fs.rmSync(lock, { force: true });
      return true;
    } catch {
      await sleep(500);
    }
  }
  return !fs.existsSync(lock);
}

/** Run a standalone ps/*.ps1 helper script; never rejects, returns {code, stdout, stderr, timedOut}. */
export function runPsScript(scriptPath, args, { timeoutMs = 120_000 } = {}) {
  return new Promise((resolve) => {
    execFile(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, ...args],
      { timeout: timeoutMs, windowsHide: true, maxBuffer: 32 * 1024 * 1024 },
      (err, stdout, stderr) => {
        resolve({
          code: err ? (typeof err.code === 'number' ? err.code : 1) : 0,
          stdout: stdout ?? '',
          stderr: stderr ?? '',
          timedOut: !!(err && err.killed),
        });
      },
    );
  });
}

export class BridgeError extends Error {
  constructor(code, message, number) {
    super(`[${code}] ${message ?? ''}`);
    this.code = code;
    this.number = number;
  }
}

// Same intent as src/bridge.ts: open/compile/backup/compact get long budgets, the
// rest fail fast. relinkCredentials gets 60s — the failure mode under test there is
// precisely a hidden modal ODBC dialog.
const OP_TIMEOUTS = {
  open: 120_000,
  compact: 180_000,
  compile: 120_000,
  backup: 120_000,
  relinkCredentials: 60_000,
};
const DEFAULT_TIMEOUT_MS = 30_000;

export class BridgeClient {
  #pending = new Map();
  #nextId = 1;
  #queue = Promise.resolve();
  #dead = false;

  constructor({ scriptPath, dbPath, stderrLog, onLog }) {
    this.scriptPath = scriptPath;
    this.dbPath = dbPath;
    this.stderrLog = stderrLog;
    this.onLog = onLog ?? (() => {});
    this.accessPid = 0;
    this.child = undefined;
  }

  get alive() {
    return !!this.child && this.child.exitCode === null && !this.#dead;
  }

  get psPid() {
    return this.child?.pid ?? 0;
  }

  start() {
    this.child = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', this.scriptPath],
      { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true },
    );
    const rl = readline.createInterface({ input: this.child.stdout });
    rl.on('line', (line) => this.#onLine(line));
    this.child.stderr.setEncoding('utf8');
    this.child.stderr.on('data', (chunk) => {
      if (this.stderrLog) fs.appendFileSync(this.stderrLog, chunk);
    });
    this.child.on('exit', () => {
      this.#dead = true;
      this.#failAllPending(new BridgeError('BRIDGE_CRASHED', 'bridge process exited'));
    });
    this.child.on('error', (err) => {
      this.#dead = true;
      this.#failAllPending(new BridgeError('BRIDGE_CRASHED', err.message));
    });
    return this;
  }

  #onLine(line) {
    let msg;
    try {
      msg = JSON.parse(line);
    } catch {
      this.onLog(`non-JSON stdout line ignored: ${line.slice(0, 200)}`);
      this.sawNonJsonStdout = true;
      return;
    }
    const entry = msg.id !== undefined && msg.id !== null ? this.#pending.get(msg.id) : undefined;
    if (!entry) return;
    this.#pending.delete(msg.id);
    clearTimeout(entry.timer);
    if (msg.ok) entry.resolve(msg.data);
    else entry.reject(new BridgeError(msg.error?.code ?? 'COM_ERROR', msg.error?.message, msg.error?.number));
  }

  #failAllPending(err) {
    for (const [, entry] of this.#pending) {
      clearTimeout(entry.timer);
      entry.reject(err);
    }
    this.#pending.clear();
  }

  /** Send one request; requests are strictly serialized (Access COM is not reentrant). */
  request(op, args = {}, timeoutMs) {
    const run = () => this.#send(op, args, timeoutMs ?? OP_TIMEOUTS[op] ?? DEFAULT_TIMEOUT_MS);
    const p = this.#queue.then(run, run);
    this.#queue = p.then(
      () => undefined,
      () => undefined,
    );
    return p;
  }

  #send(op, args, timeoutMs) {
    return new Promise((resolve, reject) => {
      if (!this.alive) return reject(new BridgeError('BRIDGE_CRASHED', 'bridge process is not running'));
      const id = this.#nextId++;
      const timer = setTimeout(() => {
        this.#pending.delete(id);
        // A hang almost always means a hidden modal dialog in MSACCESS — kill the ladder.
        void this.killLadder(`timeout ${timeoutMs}ms on op '${op}'`);
        reject(new BridgeError('BRIDGE_TIMEOUT', `op '${op}' timed out after ${timeoutMs}ms`));
      }, timeoutMs);
      this.#pending.set(id, { resolve, reject, timer });
      try {
        this.child.stdin.write(JSON.stringify({ id, op, args }) + '\n', 'utf8');
      } catch (err) {
        this.#pending.delete(id);
        clearTimeout(timer);
        reject(new BridgeError('BRIDGE_CRASHED', err.message));
      }
    });
  }

  async open({ bypassStartup = true, visibleOperations = false } = {}) {
    const data = await this.request('open', { path: this.dbPath, bypassStartup, visibleOperations });
    this.accessPid = data.accessPid ?? 0;
    return data;
  }

  /** child kill -> taskkill powershell -> taskkill MSACCESS -> delete .laccdb. */
  async killLadder(reason) {
    this.onLog(`kill ladder: ${reason}`);
    this.#dead = true;
    try {
      this.child?.kill();
    } catch {}
    await waitFor(() => !this.child || this.child.exitCode !== null, 2_000);
    if (this.child && this.child.exitCode === null) await taskkill(this.child.pid, { tree: true });
    if (this.accessPid && pidAlive(this.accessPid)) {
      await taskkill(this.accessPid, { tree: true });
      await waitFor(() => !pidAlive(this.accessPid), 5_000);
    }
    await removeLaccdb(this.dbPath);
  }

  /**
   * Graceful shutdown, observing (not forcing) cleanup first: returns what happened
   * on its own, then force-cleans any residue so the next test starts clean.
   */
  async shutdown() {
    const result = { cleanQuit: false, accessExitedItself: false, laccdbRemovedItself: false };
    const accessPid = this.accessPid;
    if (this.alive) {
      try {
        await this.request('quit', {}, 10_000);
        result.cleanQuit = true;
      } catch {}
      await waitFor(() => !this.alive, 5_000);
    }
    result.accessExitedItself = !accessPid || (await waitFor(() => !pidAlive(accessPid), 10_000));
    result.laccdbRemovedItself = await waitFor(() => !fs.existsSync(laccdbPath(this.dbPath)), 5_000, 500);
    if (!result.accessExitedItself || this.alive) await this.killLadder('residue after graceful quit');
    else if (!result.laccdbRemovedItself) await removeLaccdb(this.dbPath);
    return result;
  }
}
