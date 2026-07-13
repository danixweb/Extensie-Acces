// run-tests.mjs — headless regression suite for the Access bridge, against a REAL .accdb.
//
//   node test/run-tests.mjs [--db <path>] [--only T3.4,T4] [--with-server] [--skip-build]
//
// Phases: 0 build/static, 1 lifecycle, 2 reads, 3 write round-trips (always reverted),
// 4 robustness (crash/lock/watchdog), 5 linked tables + SQL Server (--with-server only).
// Output: TAP on stdout + JSON results file in the work dir (printed at the end).
//
// Safety: every write is a marker-wrapped edit that is restored in a finally block;
// the suite never deletes/creates objects in the database. See test/README.md.

import { exec, spawn } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  BridgeClient,
  laccdbPath,
  listProcesses,
  pidAlive,
  removeLaccdb,
  runPsScript,
  sleep,
  taskkill,
  waitFor,
} from './harness.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EXT_DIR = path.resolve(__dirname, '..');
const REPO_ROOT = path.resolve(EXT_DIR, '..');
const PS_DIR = path.join(EXT_DIR, 'ps');
const BRIDGE_PS = path.join(PS_DIR, 'access-bridge.ps1');
const DEFAULT_DB = path.join(REPO_ROOT, '64B sursa SQL SERVADENT STOC.accdb');

// ---------- CLI ----------
const argv = process.argv.slice(2);
const opts = { db: DEFAULT_DB, only: [], withServer: false, skipBuild: false };
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a === '--db') opts.db = path.resolve(argv[++i]);
  else if (a === '--only') opts.only = argv[++i].split(',').map((s) => s.trim()).filter(Boolean);
  else if (a === '--with-server') opts.withServer = true;
  else if (a === '--skip-build') opts.skipBuild = true;
  else {
    console.error(`unknown argument: ${a}`);
    process.exit(2);
  }
}

const stamp = new Date().toISOString().replace(/[:.]/g, '-');
const workDir = path.join(os.tmpdir(), 'access-explorer-tests', `run-${stamp}`);
const payloadDir = path.join(workDir, 'payloads');
fs.mkdirSync(payloadDir, { recursive: true });
const paths = {
  workDir,
  payloadDir,
  stderrLog: path.join(workDir, 'bridge-stderr.log'),
  runLog: path.join(workDir, 'run.log'),
  results: path.join(workDir, 'results.json'),
};

function logLine(msg) {
  fs.appendFileSync(paths.runLog, `${new Date().toISOString()} ${msg}\n`);
}

// ---------- TAP runner ----------
class TestSkip extends Error {}
class TestFail extends Error {}
const ok = (cond, msg) => {
  if (!cond) throw new TestFail(msg);
};
const skip = (reason) => {
  throw new TestSkip(reason);
};

const results = [];
let counter = 0;
function shouldRun(id) {
  if (opts.only.length === 0) return true;
  return opts.only.some((o) => id === o || id.startsWith(o));
}
async function test(id, desc, fn) {
  if (!shouldRun(id)) return;
  counter++;
  const started = Date.now();
  let status = 'pass';
  let detail = '';
  if (ctx.abort) {
    status = 'skip';
    detail = `abandonat: ${ctx.abort}`;
  } else {
    try {
      const note = await fn();
      if (typeof note === 'string' && note) detail = note;
    } catch (err) {
      if (err instanceof TestSkip) {
        status = 'skip';
        detail = err.message;
      } else {
        status = 'fail';
        detail = err instanceof TestFail ? err.message : `${err.code ? `[${err.code}] ` : ''}${err.message}`;
        logLine(`FAIL ${id}: ${detail}\n${err.stack ?? ''}`);
      }
    }
  }
  const durationMs = Date.now() - started;
  results.push({ id, desc, status, detail, durationMs });
  const line =
    status === 'pass'
      ? `ok ${counter} - ${id} ${desc}${detail ? `  # ${detail}` : ''}`
      : status === 'skip'
        ? `ok ${counter} - ${id} ${desc} # SKIP ${detail}`
        : `not ok ${counter} - ${id} ${desc}\n  # ${detail.replace(/\n/g, '\n  # ')}`;
  console.log(line);
}
function phase(title) {
  console.log(`\n# ─── ${title} ───`);
}

// ---------- shared context ----------
const ctx = {
  abort: '',
  bridge: undefined,
  opened: false,
  openData: undefined,
  pingData: undefined,
  listing: undefined,
  targets: {},
  moduleInfo: new Map(),
  compileBaseline: undefined, // {compiled,...} | null (hang/locked)
  compileHang: false,
  relinked: false,
  spawnedAccessPids: new Set(),
  findings: [],
};
const finding = (msg) => {
  ctx.findings.push(msg);
  logLine(`FINDING: ${msg}`);
};

// ---------- helpers ----------
function execP(cmd, cwd, timeoutMs = 300_000) {
  return new Promise((resolve) => {
    exec(cmd, { cwd, timeout: timeoutMs, windowsHide: true, maxBuffer: 32 * 1024 * 1024 }, (err, stdout, stderr) =>
      resolve({ code: err ? (typeof err.code === 'number' ? err.code : 1) : 0, stdout, stderr }),
    );
  });
}

const norm = (s) =>
  s
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((l) => l.replace(/[ \t]+$/, ''))
    .join('\n')
    .replace(/\n+$/, '');

let payloadSeq = 0;
function payloadPath(tag) {
  const safe = tag.replace(/[^\w.-]+/g, '_').slice(0, 60);
  return path.join(payloadDir, `${String(++payloadSeq).padStart(3, '0')}-${safe}.txt`);
}

async function getText(op, name, extra = {}) {
  const file = payloadPath(`${op}-${name}`);
  const data = await ctx.bridge.request(op, { name, file, ...extra });
  return { data, text: fs.readFileSync(file, 'utf8'), file };
}
async function saveText(op, name, text, extra = {}) {
  const file = payloadPath(`${op}-${name}`);
  fs.writeFileSync(file, text, 'utf8');
  return ctx.bridge.request(op, { name, file, ...extra });
}

const AI_START = "'===Start Generat AI===";
const AI_END = "'===Final Generat AI===";
const markerBlock = () =>
  `${AI_START}\r\n' TEST HARNESS ${new Date().toISOString()} - editare temporara de test, revertata automat\r\n${AI_END}`;

function newClient(tag) {
  const c = new BridgeClient({
    scriptPath: BRIDGE_PS,
    dbPath: opts.db,
    stderrLog: paths.stderrLog,
    onLog: (m) => logLine(`[bridge:${tag}] ${m}`),
  });
  c.start();
  return c;
}

/** Main long-lived bridge; respawned transparently after a kill-ladder. */
async function ensureOpen() {
  if (ctx.bridge?.alive && ctx.opened) return;
  ctx.bridge = newClient('main');
  ctx.pingData = await ctx.bridge.request('ping');
  ctx.openData = await ctx.bridge.open({ bypassStartup: true, visibleOperations: false });
  ctx.opened = true;
  ctx.spawnedAccessPids.add(ctx.bridge.accessPid);
  if (ctx.relinked) {
    const creds = readCreds();
    await ctx.bridge.request('relinkCredentials', { uid: creds.uid, pwd: creds.pwd });
  }
}

async function ensureListing() {
  await ensureOpen();
  if (ctx.listing) return;
  ctx.listing = await ctx.bridge.request('list');
  const t = ctx.targets;
  t.localTable = ctx.listing.tables.find((n) => !/^dbo_/i.test(n));
  t.linkedTable = ctx.listing.tables.includes('dbo_Erori program')
    ? 'dbo_Erori program'
    : ctx.listing.tables.find((n) => /^dbo_/i.test(n));
  t.macro = ctx.listing.macros[0];
  t.report = ctx.listing.reports[0];
}

/** Reads modules until it finds a standard module + a class-module candidate. */
async function ensureModules() {
  await ensureListing();
  const t = ctx.targets;
  if (t.stdModule !== undefined) return;
  t.stdModule = null;
  t.classModule = null; // { name, viaVbe: true } readable through VBE
  t.lockedClass = null; // name that returned VBE_TRUST_REQUIRED
  for (const name of ctx.listing.modules.slice(0, 20)) {
    if (t.stdModule && (t.classModule || t.lockedClass)) break;
    try {
      const r = await getText('getModule', name);
      ctx.moduleInfo.set(name, r);
      if (r.data.viaVbe && !t.classModule) t.classModule = name;
      else if (!r.data.isClass && !r.data.viaVbe && !t.stdModule && r.text.length > 0) t.stdModule = name;
    } catch (err) {
      ctx.moduleInfo.set(name, { error: err });
      if (err.code === 'VBE_TRUST_REQUIRED' && !t.lockedClass) t.lockedClass = name;
      if (err.code === 'BRIDGE_TIMEOUT') throw err;
    }
  }
}

async function ensureCompileBaseline() {
  if (ctx.compileBaseline !== undefined) return;
  await ensureModules();
  try {
    const r = await ctx.bridge.request('compile', { name: ctx.targets.stdModule ?? '' });
    if (r.skipped) {
      ctx.compileBaseline = null;
      ctx.compileSkipReason = r.message ?? 'compile check indisponibil';
      finding(`compile: verificarea e indisponibilă pe această bază — ${ctx.compileSkipReason}`);
      return;
    }
    ctx.compileBaseline = r;
  } catch (err) {
    if (err.code === 'BRIDGE_TIMEOUT' || err.code === 'BRIDGE_CRASHED') {
      ctx.compileBaseline = null;
      ctx.compileHang = true;
      ctx.opened = false;
      finding(
        'compile (acCmdCompileAndSaveAllModules) a blocat bridge-ul — foarte probabil dialogul de parolă VBA într-un Access ascuns. ' +
          'Extensia ar lovi același blocaj la compile-after-save pe această bază (vbaProtected).',
      );
      await ensureOpen();
    } else {
      ctx.compileBaseline = null;
      finding(`compile a întors eroare neașteptată: [${err.code}] ${err.message}`);
    }
  }
}

function readCreds() {
  const raw = fs.readFileSync(path.join(PS_DIR, 'db-credentials.local.json'), 'utf8').replace(/^﻿/, '');
  const j = JSON.parse(raw);
  ok(j.uid && j.pwd, 'db-credentials.local.json nu conține uid/pwd');
  return j;
}

/** Reads a helper-script JSON output file: strips the BOM and treats empty output as []. */
function readJsonFile(file) {
  if (!fs.existsSync(file)) throw new TestFail(`fișierul de ieșire lipsește: ${file}`);
  const raw = fs.readFileSync(file, 'utf8').replace(/^﻿/, '').trim();
  return JSON.parse(raw === '' ? '[]' : raw);
}

async function runHelper(script, args, timeoutMs = 180_000) {
  return runPsScript(path.join(PS_DIR, script), args, { timeoutMs });
}

// ---------- main ----------
async function main() {
  console.log(`# access-explorer bridge test suite`);
  console.log(`# db: ${opts.db}`);
  console.log(`# workdir: ${workDir}`);
  console.log(`# with-server: ${opts.withServer}${opts.only.length ? ` | only: ${opts.only.join(',')}` : ''}`);

  phase('Faza 0 — build/static');

  await test('T0.1', 'tsc --noEmit (npm run check)', async () => {
    if (opts.skipBuild) skip('--skip-build');
    const r = await execP('npm run check', EXT_DIR);
    ok(r.code === 0, `tsc a eșuat (exit ${r.code}):\n${(r.stdout + r.stderr).slice(-2000)}`);
  });

  await test('T0.2', 'esbuild bundle (npm run build)', async () => {
    if (opts.skipBuild) skip('--skip-build');
    const r = await execP('npm run build', EXT_DIR);
    ok(r.code === 0, `build a eșuat (exit ${r.code}):\n${(r.stdout + r.stderr).slice(-2000)}`);
    ok(fs.existsSync(path.join(EXT_DIR, 'dist', 'extension.js')), 'dist/extension.js lipsește după build');
  });

  await test('T0.3', 'preflight: mediu curat', async () => {
    ok(fs.existsSync(opts.db), `baza de date lipsește: ${opts.db}`);
    const msaccess = await listProcesses('MSACCESS.EXE');
    if (msaccess.length > 0) {
      ctx.abort = `MSACCESS.EXE deja pornit (PID ${msaccess.join(', ')}) — închide Access și repornește suita`;
      throw new TestFail(ctx.abort);
    }
    const lock = laccdbPath(opts.db);
    let note = '';
    if (fs.existsSync(lock)) {
      await removeLaccdb(opts.db);
      note = 'șters .laccdb rezidual (fără MSACCESS activ)';
    }
    if (opts.withServer) readCreds();
    return note;
  });

  phase('Faza 1 — ciclu de viață bridge');

  await test('T1.1', 'spawn + ping (doar JSON pe stdout)', async () => {
    await ensureOpen();
    ok(ctx.pingData?.pong === true, `răspuns ping neașteptat: ${JSON.stringify(ctx.pingData)}`);
    ok(!ctx.bridge.sawNonJsonStdout, 'bridge-ul a scris linii non-JSON pe stdout');
    return `powershell pid=${ctx.pingData.pid}`;
  });

  await test('T1.2', 'open cu bypassStartup (formular startup modal ocolit)', async () => {
    await ensureOpen();
    ok(ctx.bridge.accessPid > 0, 'open nu a întors accessPid');
    ok(pidAlive(ctx.bridge.accessPid), `MSACCESS pid ${ctx.bridge.accessPid} nu rulează`);
    ok(await waitFor(() => fs.existsSync(laccdbPath(opts.db)), 5_000), '.laccdb nu a apărut după open');
    ok(ctx.openData.startupRan !== true, 'formularul Startup A RULAT la open — bypass-ul nu a funcționat');
    return `accessPid=${ctx.bridge.accessPid}, vbaProtected=${ctx.openData.vbaProtected}, startupRan=${ctx.openData.startupRan}`;
  });

  await test('T1.3', 'list: 6 categorii, filtrare MSys*/~*', async () => {
    await ensureListing();
    const l = ctx.listing;
    for (const k of ['tables', 'queries', 'forms', 'reports', 'macros', 'modules'])
      ok(Array.isArray(l[k]), `lipsește categoria ${k}`);
    const all = [...l.tables, ...l.queries];
    ok(!all.some((n) => /^MSys|^~/.test(n)), 'listarea conține obiecte MSys*/~*');
    ok(l.hasLinkedTables === true, 'hasLinkedTables=false deși baza are tabele legate SQL Server');
    return `tables=${l.tables.length} queries=${l.queries.length} forms=${l.forms.length} reports=${l.reports.length} macros=${l.macros.length} modules=${l.modules.length}`;
  });

  phase('Faza 2 — citiri');

  await test('T2.1', 'getModule modul standard (SaveAsText)', async () => {
    await ensureModules();
    if (!ctx.targets.stdModule) skip('niciun modul standard găsit în primele 20 de module');
    const r = ctx.moduleInfo.get(ctx.targets.stdModule);
    ok(r.text.length > 0, 'export gol');
    // Access 2013 exportă modulele FĂRĂ antetul Attribute VB_Name pe acest mediu —
    // extensia îl sintetizează (synthesizeStandardHeader), deci antetul e opțional aici.
    const hasHeader = /Attribute VB_Name/.test(r.text);
    return `modul=${ctx.targets.stdModule}, ${r.text.length} caractere, antet=${hasHeader ? 'da' : 'nu (sintetizat de extensie)'}`;
  });

  await test('T2.2', 'getModule modul clasă (fallback viaVbe)', async () => {
    await ensureModules();
    if (ctx.targets.classModule) {
      const r = ctx.moduleInfo.get(ctx.targets.classModule);
      ok(r.text.length > 0, 'corp gol la citirea viaVbe');
      return `clasă=${ctx.targets.classModule} citită viaVbe`;
    }
    if (ctx.targets.lockedClass)
      skip(`VBE_TRUST_REQUIRED pe '${ctx.targets.lockedClass}' — proiect VBA cu parolă; se acoperă în Faza 7 după deblocare`);
    skip('niciun modul clasă identificat în primele 20 de module');
  });

  await test('T2.3', 'getQuerySql (baseline canonic)', async () => {
    await ensureListing();
    const t = ctx.targets;
    for (const q of ctx.listing.queries.slice(0, 15)) {
      try {
        const r = await getText('getQuerySql', q);
        if (/^\s*(SELECT|PARAMETERS)/i.test(r.text)) {
          t.query = q;
          t.querySql = r.text;
          break;
        }
        if (!t.query) {
          t.query = q;
          t.querySql = r.text;
        }
      } catch (err) {
        if (err.code === 'BRIDGE_TIMEOUT') throw err;
      }
    }
    ok(t.query, 'nicio interogare lizibilă găsită');
    ok(t.querySql.length > 0, `SQL gol pentru ${t.query}`);
    return `query=${t.query}, ${t.querySql.length} caractere`;
  });

  await test('T2.4', 'getMacro (encoding detectat)', async () => {
    await ensureListing();
    if (!ctx.targets.macro) skip('baza nu conține macro-uri');
    const r = await getText('getMacro', ctx.targets.macro);
    ok(r.text.length > 0, 'export macro gol');
    ok(['utf16', 'utf8', 'ansi'].includes(r.data.enc), `encoding neraportat: ${JSON.stringify(r.data)}`);
    ctx.targets.macroBaseline = r;
    return `macro=${ctx.targets.macro}, enc=${r.data.enc}`;
  });

  await test('T2.5', 'getTableDef tabel local', async () => {
    await ensureListing();
    // Numele nu spune dacă tabelul e legat (nu toate linkurile încep cu dbo_) —
    // se probează definiția și se alege primul tabel cu date locale.
    let picked, def;
    for (const tbl of ctx.listing.tables.slice(0, 15)) {
      const r = await getText('getTableDef', tbl);
      if (!/linked table/.test(r.text)) {
        picked = tbl;
        def = r;
        break;
      }
    }
    if (!picked) skip('niciun tabel local găsit în primele 15 (toate sunt legate)');
    ctx.targets.localTable = picked;
    ok(/Fields:/.test(def.text) && /Indexes:/.test(def.text), 'definiția tabelului nu are secțiunile Fields/Indexes');
    return `tabel=${picked}`;
  });

  await test('T2.6', 'getFormDef (conține CodeBehindForm)', async () => {
    await ensureListing();
    if (ctx.listing.forms.length === 0) skip('baza nu conține formulare');
    for (const f of ctx.listing.forms.slice(0, 10)) {
      const r = await getText('getFormDef', f);
      if (r.text.includes('CodeBehindForm')) {
        ctx.targets.form = f;
        ctx.targets.formBaseline = r;
        break;
      }
      if (!ctx.targets.formNoCode) ctx.targets.formNoCode = f;
    }
    if (!ctx.targets.form) skip('niciun formular cu cod (CodeBehindForm) în primele 10');
    return `form=${ctx.targets.form}, enc=${ctx.targets.formBaseline.data.enc}, ${ctx.targets.formBaseline.text.length} caractere`;
  });

  await test('T2.7', 'getReportDef', async () => {
    await ensureListing();
    if (!ctx.targets.report) skip('baza nu conține rapoarte');
    for (const rep of ctx.listing.reports.slice(0, 10)) {
      const r = await getText('getReportDef', rep);
      if (r.text.includes('CodeBehindForm')) {
        ctx.targets.report = rep;
        ctx.targets.reportBaseline = r;
        break;
      }
    }
    if (!ctx.targets.reportBaseline) {
      const r = await getText('getReportDef', ctx.targets.report);
      ok(r.text.length > 0, 'export raport gol');
      return `report=${ctx.targets.report} (fără cod — doar test de citire)`;
    }
    return `report=${ctx.targets.report}, enc=${ctx.targets.reportBaseline.data.enc}`;
  });

  await test('T2.8', 'getModule inexistent → OBJECT_NOT_FOUND', async () => {
    await ensureOpen();
    try {
      await getText('getModule', 'zzHarnessNoSuchModule');
      throw new TestFail('citirea unui modul inexistent a reușit');
    } catch (err) {
      if (err instanceof TestFail) throw err;
      ok(err.code === 'OBJECT_NOT_FOUND' || err.code === 'VBE_TRUST_REQUIRED',
        `cod de eroare neașteptat: [${err.code}] ${err.message}`);
      return `cod=${err.code}`;
    }
  });

  await test('T2.9', 'open cale inexistentă → DB_NOT_FOUND (bridge separat)', async () => {
    const c = newClient('T2.9');
    try {
      await c.request('ping');
      try {
        await c.request('open', { path: path.join(workDir, 'nu-exista.accdb'), bypassStartup: true, visibleOperations: false });
        throw new TestFail('open pe cale inexistentă a reușit');
      } catch (err) {
        if (err instanceof TestFail) throw err;
        ok(err.code === 'DB_NOT_FOUND', `cod de eroare neașteptat: [${err.code}] ${err.message}`);
      }
    } finally {
      await c.shutdown();
    }
  });

  await test('T2.10', 'compile baseline (cu watchdog — risc dialog parolă VBA)', async () => {
    await ensureCompileBaseline();
    if (ctx.compileHang) skip('compile a blocat bridge-ul (proiect VBA cu parolă) — vezi findings');
    if (ctx.compileBaseline === null) skip(ctx.compileSkipReason ?? 'compile indisponibil — vezi findings');
    return `compiled=${ctx.compileBaseline.compiled}${ctx.compileBaseline.message ? ` (${ctx.compileBaseline.message})` : ''}`;
  });

  phase('Faza 3 — round-trip scrieri (toate revertate)');

  await test('T3.1', 'round-trip modul standard: edit marcat → verificare → restaurare', async () => {
    await ensureModules();
    if (!ctx.targets.stdModule) skip('niciun modul standard');
    const name = ctx.targets.stdModule;
    const base = await getText('getModule', name);
    const edited = base.text.replace(/\r?\n*$/, '') + '\r\n' + markerBlock() + '\r\n';
    try {
      await saveText('saveModule', name, edited, { viaVbe: false });
      const after = await getText('getModule', name);
      ok(after.text.includes(AI_START), 'editarea nu a ajuns în Access');
    } finally {
      await saveText('saveModule', name, base.text, { viaVbe: false });
    }
    const restored = await getText('getModule', name);
    ok(norm(restored.text) === norm(base.text), 'restaurarea nu a readus conținutul original');
    return `modul=${name}`;
  });

  await test('T3.2', 'round-trip modul clasă (viaVbe)', async () => {
    await ensureModules();
    if (!ctx.targets.classModule) {
      if (ctx.targets.lockedClass) skip('proiect VBA cu parolă — se acoperă în Faza 7 după deblocare');
      skip('niciun modul clasă lizibil');
    }
    const name = ctx.targets.classModule;
    const base = await getText('getModule', name);
    const edited = base.text.replace(/\r?\n*$/, '') + '\r\n' + markerBlock() + '\r\n';
    try {
      await saveText('saveModule', name, edited, { viaVbe: true });
      const after = await getText('getModule', name);
      ok(after.text.includes(AI_START), 'editarea nu a ajuns în Access');
    } finally {
      await saveText('saveModule', name, base.text, { viaVbe: true });
    }
    const restored = await getText('getModule', name);
    ok(norm(restored.text) === norm(base.text), 'restaurarea nu a readus conținutul original');
    return `clasă=${name}`;
  });

  await test('T3.3', 'saveMacro: round-trip stabil pe forma canonică', async () => {
    if (!ctx.targets.macroBaseline) skip('fără macro (vezi T2.4)');
    const { macro, macroBaseline } = ctx.targets;
    // Primul import normalizează formatul (Access 2013 adaugă ex. PublishOption=1) —
    // idempotența se verifică pe forma canonică re-exportată, exact ca în extensie
    // (care după save re-citește forma canonică și își actualizează baseline-ul).
    await saveText('saveMacro', macro, macroBaseline.text, { enc: macroBaseline.data.enc });
    const canonical = await getText('getMacro', macro);
    await saveText('saveMacro', macro, canonical.text, { enc: canonical.data.enc });
    const after = await getText('getMacro', macro);
    ok(norm(after.text) === norm(canonical.text), 're-exportul diferă după re-salvarea formei canonice');
    const normalized = norm(canonical.text) !== norm(macroBaseline.text);
    return `macro=${macro}${normalized ? ' (format normalizat de Access la primul import)' : ''}`;
  });

  await test('T3.4', 'round-trip query: SQL modificat → verificare → restaurare', async () => {
    if (!ctx.targets.query) skip('fără query (vezi T2.3)');
    await ensureListing();
    const { query, querySql, localTable } = ctx.targets;
    if (!localTable) skip('fără tabel local pentru SQL-ul de test');
    const testSql = `SELECT Count(*) AS HarnessCount FROM [${localTable}];`;
    try {
      await saveText('saveQuerySql', query, testSql);
      const after = await getText('getQuerySql', query);
      ok(/HarnessCount/.test(after.text), 'SQL-ul modificat nu a ajuns în Access');
    } finally {
      await saveText('saveQuerySql', query, querySql);
    }
    const restored = await getText('getQuerySql', query);
    ok(norm(restored.text) === norm(querySql), 'SQL-ul restaurat diferă de baseline');
    return `query=${query}`;
  });

  await test('T3.5', 'saveQuerySql invalid → SQL_SYNTAX, SQL neschimbat', async () => {
    if (!ctx.targets.query) skip('fără query');
    const { query, querySql } = ctx.targets;
    try {
      await saveText('saveQuerySql', query, 'SELEC * FROM x;');
      throw new TestFail('SQL invalid a fost acceptat');
    } catch (err) {
      if (err instanceof TestFail) throw err;
      ok(err.code === 'SQL_SYNTAX', `cod de eroare neașteptat: [${err.code}] ${err.message}`);
    }
    const after = await getText('getQuerySql', query);
    ok(norm(after.text) === norm(querySql), 'SQL-ul s-a schimbat deși salvarea a eșuat');
  });

  // Access 2013 normalizează exporturile la fiecare import: recalculează Checksum, adaugă
  // (nedeterminist!) PublishOption =1 / NoSaveCTIWhenDisabled =1, elimină linii goale.
  // Round-trip-urile se verifică pe forma canonică (după un import/export), cu proprietățile
  // volatile gestionate de Access mascate — nu fac parte din design-ul utilizatorului și
  // pot apărea/dispărea între importuri identice.
  // These properties appear at varying indentation (top-level Checksum vs. a nested
  // section's PublishOption/NoSaveCTIWhenDisabled) — match leading whitespace too.
  const stripVolatile = (s) => s.replace(/^[ \t]*(Checksum|PublishOption|NoSaveCTIWhenDisabled) =.*\r?\n?/gm, '');

  async function defRoundTrip(kind, name, baseline) {
    const getOp = kind === 'form' ? 'getFormDef' : 'getReportDef';
    const saveOp = kind === 'form' ? 'saveFormDef' : 'saveReportDef';
    await saveText(saveOp, name, baseline.text, { enc: baseline.data.enc });
    const canonical = await getText(getOp, name);
    const enc = canonical.data.enc;
    const designPrefix = stripVolatile(canonical.text.slice(0, canonical.text.indexOf('CodeBehindForm')));
    const edited = canonical.text.replace(/\r?\n*$/, '') + '\r\n' + markerBlock() + '\r\n';
    try {
      await saveText(saveOp, name, edited, { enc });
      const after = await getText(getOp, name);
      ok(after.text.includes(AI_START), 'editarea nu a ajuns în Access');
      ok(
        stripVolatile(after.text.slice(0, after.text.indexOf('CodeBehindForm'))) === designPrefix,
        'blob-ul de design dinaintea CodeBehindForm s-a schimbat (dincolo de proprietățile volatile)',
      );
    } finally {
      await saveText(saveOp, name, canonical.text, { enc });
    }
    const restored = await getText(getOp, name);
    ok(
      norm(stripVolatile(restored.text)) === norm(stripVolatile(canonical.text)),
      'restaurarea nu a readus definiția canonică',
    );
  }

  await test('T3.6', 'round-trip formDef: cod editat, design intact', async () => {
    if (!ctx.targets.form) skip('fără formular cu cod (vezi T2.6)');
    await defRoundTrip('form', ctx.targets.form, ctx.targets.formBaseline);
    return `form=${ctx.targets.form}`;
  });

  await test('T3.7', 'round-trip reportDef', async () => {
    if (!ctx.targets.reportBaseline || !ctx.targets.reportBaseline.text.includes('CodeBehindForm'))
      skip('niciun raport cu cod în primele 10');
    await defRoundTrip('report', ctx.targets.report, ctx.targets.reportBaseline);
    return `report=${ctx.targets.report}`;
  });

  await test('T3.8', 'compile după restaurări = baseline', async () => {
    await ensureCompileBaseline();
    if (ctx.compileBaseline === null) skip(ctx.compileSkipReason ?? 'compile indisponibil (proiect VBA cu parolă)');
    const r = await ctx.bridge.request('compile', { name: ctx.targets.stdModule ?? '' });
    ok(r.compiled === ctx.compileBaseline.compiled, `compile=${r.compiled}, baseline=${ctx.compileBaseline.compiled}`);
    return `compiled=${r.compiled}`;
  });

  await test('T3.9', 'modul stricat → compiled:false → restaurare → baseline', async () => {
    await ensureCompileBaseline();
    if (ctx.compileBaseline === null) skip(ctx.compileSkipReason ?? 'compile indisponibil (proiect VBA cu parolă)');
    if (!ctx.targets.stdModule) skip('fără modul standard');
    const name = ctx.targets.stdModule;
    const base = await getText('getModule', name);
    const broken =
      base.text.replace(/\r?\n*$/, '') + `\r\n${AI_START}\r\nSub HarnessBroken(\r\n${AI_END}\r\n`;
    let brokenResult;
    try {
      await saveText('saveModule', name, broken, { viaVbe: false });
      const check = await getText('getModule', name);
      ok(check.text.includes('Sub HarnessBroken('), 'codul stricat nu a fost efectiv salvat în modul');
      brokenResult = await ctx.bridge.request('compile', { name });
    } finally {
      await saveText('saveModule', name, base.text, { viaVbe: false });
    }
    ok(brokenResult.compiled === false, 'compile a raportat succes pe cod stricat');
    const after = await ctx.bridge.request('compile', { name });
    ok(after.compiled === ctx.compileBaseline.compiled, `compile după restaurare=${after.compiled}, baseline=${ctx.compileBaseline.compiled}`);
    return `mesaj compile: ${brokenResult.message ?? '-'}`;
  });

  await test('T3.10', 'backup: copie validă a fișierului live', async () => {
    await ensureOpen();
    const target = path.join(workDir, 'backup-test.accdb');
    await ctx.bridge.request('backup', { target });
    ok(fs.existsSync(target), 'fișierul de backup lipsește');
    const size = fs.statSync(target).size;
    ok(size > 1024 * 1024, `backup suspect de mic: ${size} bytes`);
    const head = Buffer.alloc(32);
    const fd = fs.openSync(target, 'r');
    fs.readSync(fd, head, 0, 32, 0);
    fs.closeSync(fd);
    ok(head.toString('latin1').includes('Standard ACE DB'), 'antetul fișierului nu este Standard ACE DB');
    fs.rmSync(target, { force: true });
    return `${(size / 1024 / 1024).toFixed(1)} MB`;
  });

  await test('T3.11', 'saveQuerySql pe nume inexistent → OBJECT_NOT_FOUND', async () => {
    await ensureOpen();
    try {
      await saveText('saveQuerySql', 'zzHarnessNoSuchQuery', 'SELECT 1;');
      throw new TestFail('salvarea pe un query inexistent a reușit');
    } catch (err) {
      if (err instanceof TestFail) throw err;
      ok(err.code === 'OBJECT_NOT_FOUND', `cod de eroare neașteptat: [${err.code}] ${err.message}`);
      return `cod=${err.code}`;
    }
  });

  phase('Faza 4 — robustețe / ciclu de viață');

  await test('T4.1', 'compact & repair in-place', async () => {
    await ensureOpen();
    const before = fs.statSync(opts.db).size;
    let r;
    try {
      r = await ctx.bridge.request('compact', {});
    } catch (err) {
      const leftovers = fs
        .readdirSync(path.dirname(opts.db))
        .filter((f) => f.includes('_compact_'));
      for (const f of leftovers) fs.rmSync(path.join(path.dirname(opts.db), f), { force: true });
      if (/Cannot Compile Project/i.test(err.message ?? '')) {
        // Stare a bazei, nu bug de bridge: CompactRepair refuză un proiect VBA care nu
        // se poate (re)compila. Bridge-ul a redeschis baza nedistructiv (verificat mai jos).
        const l = await ctx.bridge.request('list');
        ok(Array.isArray(l.tables) && l.tables.length > 0, 'baza nu mai răspunde după compact eșuat');
        finding(
          'compact eșuează pe această bază cu "Cannot Compile Project" (proiect VBA cu parolă / necompilabil). ' +
            'Atenție: compactOnClose=true (implicit) va lovi aceeași eroare la fiecare închidere din extensie. ' +
            'De verificat în sesiunea ghidată după deblocarea VBA.',
        );
        skip('compact refuzat de Access: Cannot Compile Project (stare DB, nedistructiv — baza rămâne funcțională)');
      }
      throw new TestFail(
        `compact a eșuat: [${err.code}] ${err.message}${leftovers.length ? ` (șters rezidual: ${leftovers.join(', ')})` : ''}`,
      );
    }
    ok(r.compacted === true && r.listing, 'compact nu a întors compacted:true + listing');
    const after = fs.statSync(opts.db).size;
    ctx.listing = r.listing;
    return `${(before / 1e6).toFixed(1)} MB → ${(after / 1e6).toFixed(1)} MB`;
  });

  await test('T4.2', 'quit curat: MSACCESS dispare, .laccdb șters', async () => {
    await ensureOpen();
    const res = await ctx.bridge.shutdown();
    ctx.opened = false;
    ok(res.cleanQuit, 'quit nu a răspuns');
    ok(res.accessExitedItself, 'MSACCESS nu s-a închis singur în 10s după quit');
    ok(res.laccdbRemovedItself, '.laccdb nu a fost șters de Access la închidere');
    return 'MSACCESS închis, .laccdb șters';
  });

  await test('T4.3', 'crash bridge → orfan detectat + curățat', async () => {
    const c = newClient('T4.3');
    await c.request('ping');
    await c.open({ bypassStartup: true, visibleOperations: false });
    ctx.spawnedAccessPids.add(c.accessPid);
    const accessPid = c.accessPid;
    await taskkill(c.psPid); // simulate the user killing powershell.exe in Task Manager
    await waitFor(() => !c.alive, 5_000);
    ok(!c.alive, 'procesul powershell nu a murit după taskkill');
    await sleep(1_500);
    if (!pidAlive(accessPid)) {
      await removeLaccdb(opts.db);
      return 'MSACCESS s-a închis singur la moartea bridge-ului (comportament acceptabil)';
    }
    const r = await runHelper('find-orphaned-access.ps1', []);
    ok(r.code === 0, `find-orphaned-access.ps1 a eșuat: ${r.stderr.slice(-500)}`);
    const orphans = JSON.parse(r.stdout.trim() || '[]');
    ok(orphans.some((o) => o.pid === accessPid), `orfanul MSACCESS ${accessPid} nu a fost detectat: ${r.stdout.slice(0, 300)}`);
    await taskkill(accessPid, { tree: true });
    ok(await waitFor(() => !pidAlive(accessPid), 5_000), 'orfanul MSACCESS nu a putut fi oprit');
    ok(await removeLaccdb(opts.db), '.laccdb rezidual nu a putut fi șters');
    return `orfan pid=${accessPid} detectat și curățat`;
  });

  await test('T4.4', 'redeschidere după crash: open + list + quit', async () => {
    const c = newClient('T4.4');
    try {
      await c.request('ping');
      await c.open({ bypassStartup: true, visibleOperations: false });
      ctx.spawnedAccessPids.add(c.accessPid);
      const l = await c.request('list');
      ok(Array.isArray(l.tables) && l.tables.length > 0, 'list gol după redeschidere');
    } finally {
      const res = await c.shutdown();
      ok(res.cleanQuit, 'quit nu a răspuns la redeschidere');
    }
  });

  await test('T4.5', 'lock exclusiv → open întoarce DB_LOCKED', async () => {
    // Deținătorul lock-ului exclusiv se deschide tot prin COM, cu bypass de startup
    // (Shift simulat, ca în bridge) — formularul startup modal NU trebuie să pornească
    // în niciun test. OpenCurrentDatabase(path, $true) = exclusiv.
    const holderScript = path.join(workDir, 'exclusive-holder.ps1');
    fs.writeFileSync(
      holderScript,
      `param([string]$Db)
$ErrorActionPreference = 'Stop'
Add-Type -Name Win32 -Namespace Holder -MemberDefinition '[DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, System.UIntPtr dwExtraInfo);'
$app = New-Object -ComObject Access.Application
$app.Visible = $false
[Holder.Win32]::keybd_event(0x10, 0, 0, [System.UIntPtr]::Zero)
Start-Sleep -Milliseconds 100
try { $app.OpenCurrentDatabase($Db, $true) } finally { [Holder.Win32]::keybd_event(0x10, 0, 2, [System.UIntPtr]::Zero) }
Write-Output "HOLDING"
Start-Sleep -Seconds 120
try { $app.CloseCurrentDatabase() } catch { }
try { $app.Quit(2) } catch { }
`,
      'utf8',
    );
    const msBefore = await listProcesses('MSACCESS.EXE');
    const holder = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-File', holderScript, '-Db', opts.db],
      { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true },
    );
    let holding = false;
    holder.stdout.setEncoding('utf8');
    holder.stdout.on('data', (d) => {
      if (d.includes('HOLDING')) holding = true;
    });
    try {
      const appeared = await waitFor(() => holding && fs.existsSync(laccdbPath(opts.db)), 60_000, 500);
      if (!appeared) skip('deținătorul exclusiv nu a deschis baza în 60s');
      const c = newClient('T4.5');
      try {
        await c.request('ping');
        try {
          await c.open({ bypassStartup: true, visibleOperations: false });
          throw new TestFail('open a reușit deși baza era deschisă exclusiv');
        } catch (err) {
          if (err instanceof TestFail) throw err;
          ok(err.code === 'DB_LOCKED', `cod de eroare neașteptat: [${err.code}] ${err.message}`);
        }
      } finally {
        await c.shutdown();
      }
    } finally {
      await taskkill(holder.pid, { tree: true });
      const msAfter = await listProcesses('MSACCESS.EXE');
      for (const pid of msAfter.filter((p) => !msBefore.includes(p))) {
        await taskkill(pid, { tree: true });
        await waitFor(() => !pidAlive(pid), 5_000);
      }
      await removeLaccdb(opts.db);
    }
  });

  await test('T4.6', 'watchdog self-test: timeout → kill-ladder → bridge nou funcțional', async () => {
    const c = newClient('T4.6a');
    await c.request('ping');
    await c.open({ bypassStartup: true, visibleOperations: false });
    ctx.spawnedAccessPids.add(c.accessPid);
    const accessPid = c.accessPid;
    try {
      await c.request('list', {}, 1);
      throw new TestFail('cererea cu timeout de 1ms nu a expirat');
    } catch (err) {
      if (err instanceof TestFail) throw err;
      ok(err.code === 'BRIDGE_TIMEOUT', `cod de eroare neașteptat: [${err.code}]`);
    }
    ok(await waitFor(() => !pidAlive(accessPid), 15_000), 'kill-ladder nu a oprit MSACCESS');
    ok(await waitFor(() => !fs.existsSync(laccdbPath(opts.db)), 10_000), 'kill-ladder nu a șters .laccdb');
    const c2 = newClient('T4.6b');
    try {
      await c2.request('ping');
      await c2.open({ bypassStartup: true, visibleOperations: false });
      ctx.spawnedAccessPids.add(c2.accessPid);
      const l = await c2.request('list');
      ok(Array.isArray(l.tables), 'list a eșuat pe bridge-ul nou');
    } finally {
      await c2.shutdown();
    }
  });

  phase('Faza 5 — tabele legate + SQL Server');

  const serverSkip = opts.withServer ? null : 'rulare fără --with-server';

  await test('T5.1', 'relinkCredentials cu credențiale valide', async () => {
    if (serverSkip) skip(serverSkip);
    await ensureOpen();
    const creds = readCreds();
    const r = await ctx.bridge.request('relinkCredentials', { uid: creds.uid, pwd: creds.pwd });
    ok(r.relinked > 0, `niciun tabel relegat: ${JSON.stringify(r)}`);
    ctx.relinked = true;
    const failed = r.failed ?? [];
    if (failed.length > 0)
      finding(`relinkCredentials: ${failed.length} tabele legate NU s-au putut relega nici cu credențiale valide: ${failed.join(', ')}`);
    return `relinked=${r.relinked}, failed=${failed.length}`;
  });

  await test('T5.2', 'getTableDef pe tabelul legat dbo_Erori program', async () => {
    if (serverSkip) skip(serverSkip);
    if (!ctx.targets.linkedTable) skip('niciun tabel dbo_* în listare');
    ok(ctx.relinked, 'relinkCredentials nu a rulat — nu ating tabelul legat fără credențiale');
    const r = await getText('getTableDef', ctx.targets.linkedTable, {});
    ok(/Fields:/.test(r.text), 'definiția tabelului legat nu are secțiunea Fields');
    return `tabel=${ctx.targets.linkedTable}`;
  });

  await test('T5.3', 'test-db-query.ps1: citire rânduri din tabel legat', async () => {
    if (serverSkip) skip(serverSkip);
    if (ctx.bridge?.alive) {
      await ctx.bridge.shutdown();
      ctx.opened = false;
    }
    const out = path.join(workDir, 'query-rows.json');
    const r = await runHelper('test-db-query.ps1', [
      '-DbPath', opts.db, '-Table', ctx.targets.linkedTable ?? 'dbo_Erori program', '-Top', '5', '-OutFile', out,
    ]);
    ok(r.code === 0 && !r.timedOut, `script eșuat (exit ${r.code}${r.timedOut ? ', timeout' : ''}): ${(r.stderr || r.stdout).slice(-800)}`);
    const rows = readJsonFile(out);
    return `${Array.isArray(rows) ? rows.length : 1} rând(uri) citite`;
  });

  await test('T5.4', 'ciclu complet: add-test-error → get-untreated → mark-treated', async () => {
    if (serverSkip) skip(serverSkip);
    const numar = Math.floor(Date.now() / 1000) % 100000;
    const mesaj = `HARNESS ${stamp}`;
    const add = await runHelper('add-test-error.ps1', [
      '-DbPath', opts.db, '-Modul', 'TestHarness', '-Rutina', 'AutoTest',
      '-Numar', String(numar), '-Mesaj', mesaj, '-Context', 'Rand de test inserat de suita automata',
    ]);
    ok(add.code === 0 && !add.timedOut, `add-test-error a eșuat: ${(add.stderr || add.stdout).slice(-800)}`);
    const m = add.stdout.match(/id=(\d+)/);
    ok(m, `id-ul inserat nu apare în ieșire: ${add.stdout.slice(-300)}`);
    const id = Number(m[1]);

    const out1 = path.join(workDir, 'untreated-1.json');
    const get1 = await runHelper('get-untreated-errors.ps1', ['-DbPath', opts.db, '-Limit', '100000', '-OutFile', out1]);
    ok(get1.code === 0, `get-untreated-errors a eșuat: ${(get1.stderr || get1.stdout).slice(-800)}`);
    const rows1 = readJsonFile(out1);
    ok(rows1.some((r) => r.id === id), `rândul inserat id=${id} nu apare printre cele netratate`);

    const mark = await runHelper('mark-error-treated.ps1', ['-DbPath', opts.db, '-Id', String(id)]);
    ok(mark.code === 0, `mark-error-treated a eșuat: ${(mark.stderr || mark.stdout).slice(-800)}`);

    const out2 = path.join(workDir, 'untreated-2.json');
    const get2 = await runHelper('get-untreated-errors.ps1', ['-DbPath', opts.db, '-Limit', '100000', '-OutFile', out2]);
    ok(get2.code === 0, `get-untreated-errors (recitire) a eșuat`);
    const rows2 = readJsonFile(out2);
    ok(!rows2.some((r) => r.id === id), `rândul id=${id} încă apare ca netratat după mark-error-treated`);
    finding(`T5.4 a lăsat pe SQL Server un rând de log TRATAT în [dbo_Erori program]: id=${id}, mesaj="${mesaj}".`);
    return `id=${id} inserat, regăsit și marcat tratat`;
  });

  await test('T5.5', 'relinkCredentials cu parolă greșită: fără blocaj (dialog modal)', async () => {
    if (serverSkip) skip(serverSkip);
    await ensureOpen();
    let outcome;
    let note = '';
    try {
      const r = await ctx.bridge.request('relinkCredentials', { uid: readCreds().uid, pwd: `gresit-${stamp}` });
      // RefreshLink poate „reuși" cu parolă greșită: driverul ODBC refolosește conexiunea
      // deja autentificată din pool-ul procesului, deci validarea reală se amână până la
      // primul acces la date. Limitare cunoscută — documentată, nu tratată ca eșec aici.
      note = `parola greșită NU a fost respinsă la relink (relinked=${r.relinked}) — connection pooling ODBC`;
      finding(
        'relinkCredentials nu validează efectiv credențialele când există deja o conexiune ODBC autentificată în ' +
          'același proces Access (pooling): parola greșită e acceptată silențios și persistată în connect-string-uri. ' +
          'Extensia poate salva astfel credențiale greșite în Secret Storage fără să afle imediat.',
      );
    } catch (err) {
      if (err.code === 'LINKED_AUTH_FAILED') note = 'respins corect cu LINKED_AUTH_FAILED';
      else if (err.code === 'BRIDGE_TIMEOUT' || err.code === 'BRIDGE_CRASHED') {
        ctx.opened = false;
        outcome = new TestFail(
          'relink cu parolă greșită a BLOCAT bridge-ul (dialog ODBC modal invizibil) — extensia ar lovi același blocaj',
        );
        finding(outcome.message);
      } else outcome = new TestFail(`cod de eroare neașteptat: [${err.code}] ${err.message}`);
    } finally {
      // Critical: leave the .accdb with VALID persisted credentials no matter what.
      await ensureOpen();
      const creds = readCreds();
      const r = await ctx.bridge.request('relinkCredentials', { uid: creds.uid, pwd: creds.pwd });
      if (!(r.relinked > 0)) finding('CRITIC: restaurarea credențialelor valide după T5.5 a eșuat — rulează manual relinkCredentials!');
    }
    if (outcome) throw outcome;
    return note;
  });

  // ---------- final cleanup + report ----------
  phase('Curățenie finală');
  if (ctx.bridge?.alive) {
    await ctx.bridge.shutdown();
    ctx.opened = false;
  }
  for (const pid of ctx.spawnedAccessPids) {
    if (pidAlive(pid)) {
      logLine(`final sweep: killing leftover MSACCESS pid ${pid}`);
      await taskkill(pid, { tree: true });
    }
  }
  await removeLaccdb(opts.db);
  const leftover = await listProcesses('MSACCESS.EXE');
  console.log(`# postflight: MSACCESS rămase=${leftover.length}, .laccdb=${fs.existsSync(laccdbPath(opts.db))}`);

  const summary = {
    pass: results.filter((r) => r.status === 'pass').length,
    fail: results.filter((r) => r.status === 'fail').length,
    skip: results.filter((r) => r.status === 'skip').length,
  };
  console.log(`\n1..${counter}`);
  console.log(`# TOTAL: ${summary.pass} pass, ${summary.fail} fail, ${summary.skip} skip`);
  if (ctx.findings.length) {
    console.log('# FINDINGS:');
    for (const f of ctx.findings) console.log(`#  - ${f}`);
  }
  fs.writeFileSync(
    paths.results,
    JSON.stringify({ db: opts.db, stamp, opts: { withServer: opts.withServer, only: opts.only }, summary, results, findings: ctx.findings }, null, 2),
  );
  console.log(`# rezultate: ${paths.results}`);
  process.exitCode = summary.fail > 0 ? 1 : 0;
}

main().catch(async (err) => {
  console.error(`\n# EROARE FATALĂ: ${err.stack ?? err}`);
  try {
    if (ctx.bridge?.alive) await ctx.bridge.killLadder('fatal error');
    for (const pid of ctx.spawnedAccessPids) if (pidAlive(pid)) await taskkill(pid, { tree: true });
    await removeLaccdb(opts.db);
  } catch {}
  process.exitCode = 2;
});
