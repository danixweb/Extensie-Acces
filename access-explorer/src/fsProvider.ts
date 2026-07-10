import * as vscode from 'vscode';
import { BackupManager } from './backup';
import { BridgeError, describeError } from './errors';
import {
  Category,
  DatabaseRegistry,
  EDITABLE_CATEGORIES,
  EXT_FOR,
  listingFor,
  objectUri,
  OpenDatabase,
  parseUri,
  SCHEME,
} from './model';
import {
  joinModuleHeader,
  splitFormHeader,
  splitModuleHeader,
  synthesizeStandardHeader,
} from './vbaHeader';

interface Entry {
  category: Category;
  name: string;
  readonly: boolean;
  mtime: number;
  /** Text served to the editor (module body without the SaveAsText header). */
  text: string;
  /** Content snapshot used to detect external changes at save time. */
  baseline: string;
  /** Leading SaveAsText header block (Modules only), restored on write. */
  header?: string;
  /** True when the module was read through the VBE and must be written back the same way. */
  viaVbe?: boolean;
  /** Encoding of the macro export, reused on write-back. */
  macroEnc?: string;
  /** Encoding of a Form/Report code export, reused on write-back. */
  codeEnc?: string;
}

function normalize(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\s+$/, '');
}

/**
 * FileSystemProvider for the accessdb: scheme. Ctrl+S routes through writeFile,
 * which does: conflict check -> backup -> write to Access -> compile check.
 */
export class AccessFsProvider implements vscode.FileSystemProvider {
  private readonly entries = new Map<string, Entry>();
  private readonly fileChangeEmitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
  readonly onDidChangeFile = this.fileChangeEmitter.event;

  constructor(
    private readonly registry: DatabaseRegistry,
    private readonly backups: BackupManager,
  ) {}

  register(): vscode.Disposable {
    return vscode.workspace.registerFileSystemProvider(SCHEME, this, {
      isCaseSensitive: false,
    });
  }

  /**
   * Builds the URI for an object, fetching its content first when the extension
   * depends on it (module kind decides .bas vs .cls) and priming the cache.
   */
  async resolveUri(db: OpenDatabase, category: Category, name: string): Promise<vscode.Uri> {
    if (category !== 'Modules') {
      return objectUri(db.key, category, name, EXT_FOR[category]);
    }
    const { text, isClass, viaVbe } = await db.bridge.getModule(name);
    const uri = objectUri(db.key, category, name, isClass ? '.cls' : '.bas');
    const { header, body } = splitModuleHeader(text);
    this.entries.set(uri.toString(), {
      category,
      name,
      readonly: false,
      mtime: Date.now(),
      text: body,
      baseline: body,
      header,
      viaVbe,
    });
    return uri;
  }

  /** Drops cached content for a database so open editors re-read fresh exports. */
  invalidate(key: string): void {
    const events: vscode.FileChangeEvent[] = [];
    for (const [uriString, entry] of [...this.entries]) {
      if (uriString.includes(`/${key}/`)) {
        this.entries.delete(uriString);
        entry.mtime = Date.now();
        events.push({ type: vscode.FileChangeType.Changed, uri: vscode.Uri.parse(uriString) });
      }
    }
    if (events.length > 0) {
      this.fileChangeEmitter.fire(events);
    }
  }

  dropDatabase(key: string): void {
    const events: vscode.FileChangeEvent[] = [];
    for (const uriString of [...this.entries.keys()]) {
      if (uriString.includes(`/${key}/`)) {
        this.entries.delete(uriString);
        events.push({ type: vscode.FileChangeType.Deleted, uri: vscode.Uri.parse(uriString) });
      }
    }
    if (events.length > 0) {
      this.fileChangeEmitter.fire(events);
    }
  }

  // ---------- FileSystemProvider ----------

  watch(): vscode.Disposable {
    return new vscode.Disposable(() => undefined);
  }

  async stat(uri: vscode.Uri): Promise<vscode.FileStat> {
    const entry = await this.ensureEntry(uri);
    return {
      type: vscode.FileType.File,
      ctime: entry.mtime,
      mtime: entry.mtime,
      size: Buffer.byteLength(entry.text, 'utf8'),
      permissions: entry.readonly ? vscode.FilePermission.Readonly : undefined,
    };
  }

  async readFile(uri: vscode.Uri): Promise<Uint8Array> {
    const entry = await this.ensureEntry(uri);
    return Buffer.from(entry.text, 'utf8');
  }

  async writeFile(uri: vscode.Uri, content: Uint8Array): Promise<void> {
    await this.saveObject(uri, Buffer.from(content).toString('utf8'));
  }

  /**
   * Full save pipeline for an object's text: conflict check, backup, COM write, compile check,
   * re-read of the canonical form. Shared by the accessdb: FileSystemProvider.writeFile (Ctrl+S)
   * and the on-disk AI mirror, so both editing surfaces get identical safety semantics.
   */
  async saveObject(uri: vscode.Uri, newText: string): Promise<{ text: string; readonly: boolean }> {
    const { key, category, name } = parseUri(uri);
    const db = this.registry.get(key);
    if (!db) {
      throw vscode.FileSystemError.Unavailable(
        vscode.l10n.t('The database is no longer open in Access Explorer.'),
      );
    }
    if (!EDITABLE_CATEGORIES.has(category)) {
      throw vscode.FileSystemError.NoPermissions(uri);
    }
    const entry = await this.ensureEntry(uri);
    if (entry.readonly) {
      // Forms/Reports with no code-behind module (HasModule=False) have nothing to write.
      throw vscode.FileSystemError.NoPermissions(uri);
    }

    try {
      await this.checkConflict(db, entry);
      // Backup failures abort the write — that is the whole point of the backup.
      await this.backups.beforeWrite(db);

      switch (category) {
        case 'Modules': {
          if (entry.viaVbe) {
            // VBE-read modules carry no header and must not go through LoadFromText.
            await db.bridge.saveModule(name, newText, true);
          } else {
            const header = entry.header ?? synthesizeStandardHeader(name);
            await db.bridge.saveModule(name, joinModuleHeader(header, newText), false);
          }
          this.scheduleCompileCheck(db, name);
          break;
        }
        case 'Queries':
          await db.bridge.saveQuerySql(name, newText);
          break;
        case 'Macros':
          await db.bridge.saveMacro(name, newText, entry.macroEnc ?? 'utf16');
          break;
        case 'Forms':
        case 'Reports': {
          const kind = category === 'Forms' ? 'form' : 'report';
          // entry.header is the untouched design blob (Begin/End blocks) cached at load time —
          // never parsed, just spliced back verbatim ahead of the edited code.
          const joined = joinModuleHeader(entry.header ?? '', newText);
          await db.bridge.saveFormOrReportDef(kind, name, joined, entry.codeEnc ?? 'ansi');
          this.scheduleCompileCheck(db, name);
          break;
        }
      }
    } catch (err) {
      throw vscode.FileSystemError.Unavailable(describeError(err, { vbaLocked: db.vbaLocked }));
    }

    // Access may store a canonical form of what was written (e.g. DAO reformats
    // query SQL). Re-read it so the next conflict check compares against what the
    // database actually contains — and so the editor shows the stored form.
    try {
      const fresh = await this.loadEntry(db, category, name);
      entry.text = fresh.text;
      entry.baseline = fresh.text;
      entry.header = fresh.header;
      entry.viaVbe = fresh.viaVbe;
      entry.macroEnc = fresh.macroEnc ?? entry.macroEnc;
      entry.codeEnc = fresh.codeEnc ?? entry.codeEnc;
    } catch {
      entry.text = newText;
      entry.baseline = newText;
    }
    entry.mtime = Date.now();
    this.fileChangeEmitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    return { text: entry.text, readonly: entry.readonly };
  }

  /** Resolves an object's uri and current text, reusing the same cache as writeFile/resolveUri. */
  async readObject(
    db: OpenDatabase,
    category: Category,
    name: string,
  ): Promise<{ uri: vscode.Uri; text: string; readonly: boolean }> {
    const uri = await this.resolveUri(db, category, name);
    const entry = await this.ensureEntry(uri);
    return { uri, text: entry.text, readonly: entry.readonly };
  }

  /** Drops the cached entry for one object, forcing a fresh read next time it's touched. */
  invalidateOne(uri: vscode.Uri): void {
    if (this.entries.delete(uri.toString())) {
      this.fileChangeEmitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }
  }

  readDirectory(uri: vscode.Uri): [string, vscode.FileType][] {
    const parts = uri.path.split('/').filter((p) => p.length > 0);
    if (parts.length === 1) {
      const db = this.registry.get(parts[0]);
      if (!db) {
        throw vscode.FileSystemError.FileNotFound(uri);
      }
      return ['Tables', 'Queries', 'Forms', 'Reports', 'Macros', 'Modules'].map((c) => [
        c,
        vscode.FileType.Directory,
      ]);
    }
    if (parts.length === 2) {
      const db = this.registry.get(parts[0]);
      if (!db) {
        throw vscode.FileSystemError.FileNotFound(uri);
      }
      const category = parts[1] as Category;
      const ext = category === 'Modules' ? '.bas' : EXT_FOR[category];
      return listingFor(db.listing, category).map((n) => [n + ext, vscode.FileType.File]);
    }
    throw vscode.FileSystemError.FileNotFound(uri);
  }

  createDirectory(uri: vscode.Uri): void {
    throw vscode.FileSystemError.NoPermissions(uri);
  }

  delete(uri: vscode.Uri): void {
    throw vscode.FileSystemError.NoPermissions(uri);
  }

  rename(uri: vscode.Uri): void {
    throw vscode.FileSystemError.NoPermissions(uri);
  }

  // ---------- internals ----------

  private async ensureEntry(uri: vscode.Uri): Promise<Entry> {
    const cached = this.entries.get(uri.toString());
    if (cached) {
      return cached;
    }
    const { key, category, name } = parseUri(uri);
    const db = this.registry.get(key);
    if (!db) {
      throw vscode.FileSystemError.FileNotFound(uri);
    }
    try {
      const entry = await this.loadEntry(db, category, name);
      this.entries.set(uri.toString(), entry);
      return entry;
    } catch (err) {
      if (err instanceof BridgeError && err.code === 'OBJECT_NOT_FOUND') {
        throw vscode.FileSystemError.FileNotFound(uri);
      }
      throw vscode.FileSystemError.Unavailable(describeError(err, { vbaLocked: db.vbaLocked }));
    }
  }

  private async loadEntry(db: OpenDatabase, category: Category, name: string): Promise<Entry> {
    const base = { category, name, mtime: Date.now() };
    switch (category) {
      case 'Modules': {
        const { text, viaVbe } = await db.bridge.getModule(name);
        const { header, body } = splitModuleHeader(text);
        return { ...base, readonly: false, text: body, baseline: body, header, viaVbe };
      }
      case 'Queries': {
        const sql = await db.bridge.getQuerySql(name);
        return { ...base, readonly: false, text: sql, baseline: sql };
      }
      case 'Macros': {
        const { text, enc } = await db.bridge.getMacro(name);
        return { ...base, readonly: false, text, baseline: text, macroEnc: enc };
      }
      case 'Tables': {
        const { text } = await db.bridge.getReadonlyDef('table', name);
        return { ...base, readonly: true, text, baseline: text };
      }
      case 'Forms':
      case 'Reports': {
        const kind = category === 'Forms' ? 'form' : 'report';
        const { text: fullText, enc } = await db.bridge.getReadonlyDef(kind, name);
        const split = splitFormHeader(fullText);
        if (!split) {
          // No CodeBehindForm section (HasModule=False) — nothing to edit, show as-is.
          return { ...base, readonly: true, text: fullText, baseline: fullText };
        }
        return {
          ...base,
          readonly: false,
          text: split.body,
          baseline: split.body,
          header: split.header,
          codeEnc: enc,
        };
      }
    }
  }

  /**
   * Re-exports the object and compares with the snapshot taken when the editor
   * loaded it. A mismatch means someone changed it in Access since — refuse the
   * save instead of silently overwriting their work.
   */
  private async checkConflict(db: OpenDatabase, entry: Entry): Promise<void> {
    let current: string;
    switch (entry.category) {
      case 'Modules': {
        const fresh = await db.bridge.getModule(entry.name);
        const split = splitModuleHeader(fresh.text);
        entry.header = split.header; // freshest header wins
        entry.viaVbe = fresh.viaVbe;
        current = split.body;
        break;
      }
      case 'Queries':
        current = await db.bridge.getQuerySql(entry.name);
        break;
      case 'Macros':
        current = (await db.bridge.getMacro(entry.name)).text;
        break;
      case 'Forms':
      case 'Reports': {
        const kind = entry.category === 'Forms' ? 'form' : 'report';
        const fresh = await db.bridge.getReadonlyDef(kind, entry.name);
        const split = splitFormHeader(fresh.text);
        if (!split) {
          return;
        }
        entry.header = split.header; // freshest design blob wins
        entry.codeEnc = fresh.enc ?? entry.codeEnc;
        current = split.body;
        break;
      }
      default:
        return;
    }
    if (normalize(current) !== normalize(entry.baseline)) {
      throw new BridgeError('CONFLICT');
    }
  }

  /** Fire-and-forget compile check after a module save; failure is a warning, not a rollback. */
  private scheduleCompileCheck(db: OpenDatabase, moduleName: string): void {
    const cfg = vscode.workspace.getConfiguration('accessExplorer');
    if (!cfg.get<boolean>('compileAfterSave', true)) {
      return;
    }
    void db.bridge
      .compile(moduleName)
      .then((result) => {
        if (!result.compiled) {
          void vscode.window.showWarningMessage(
            vscode.l10n.t(
              'Saved, but the VBA project does not compile: {0}',
              result.message ?? vscode.l10n.t('unknown compile error'),
            ),
          );
        }
      })
      .catch((err) => {
        void vscode.window.showWarningMessage(
          vscode.l10n.t('Saved, but the compile check failed: {0}', describeError(err)),
        );
      });
  }
}
