import * as crypto from 'node:crypto';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { AccessBridge, DbListing } from './bridge';

export const SCHEME = 'accessdb';

export type Category = 'Tables' | 'Queries' | 'Forms' | 'Reports' | 'Macros' | 'Modules';

export const CATEGORIES: Category[] = ['Tables', 'Queries', 'Forms', 'Reports', 'Macros', 'Modules'];

export const EDITABLE_CATEGORIES: ReadonlySet<Category> = new Set([
  'Queries',
  'Macros',
  'Modules',
  'Forms',
  'Reports',
]);

export const EXT_FOR: Record<Exclude<Category, 'Modules'>, string> = {
  Tables: '.table.txt',
  Queries: '.sql',
  Forms: '.form.txt',
  Reports: '.report.txt',
  Macros: '.mac',
};

export const MODULE_EXTS = ['.cls', '.bas'] as const;

export interface OpenDatabase {
  key: string;
  dbPath: string;
  bridge: AccessBridge;
  listing: DbListing;
  /** True when the VBA project is password-locked (Tools > VBAProject Properties > Protection). */
  vbaLocked: boolean;
}

export function dbKeyFor(dbPath: string): string {
  return crypto.createHash('sha1').update(dbPath.toLowerCase()).digest('hex').slice(0, 10);
}

export function listingFor(listing: DbListing, category: Category): string[] {
  switch (category) {
    case 'Tables': return listing.tables;
    case 'Queries': return listing.queries;
    case 'Forms': return listing.forms;
    case 'Reports': return listing.reports;
    case 'Macros': return listing.macros;
    case 'Modules': return listing.modules;
  }
}

// Access object names may contain characters that break URI paths; encode just those.
export function encodeName(name: string): string {
  return name.replace(/[%/\\#?]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'));
}

export function decodeName(segment: string): string {
  return segment.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

/**
 * Extension for a filename within a category, matched against the known suffix rather than
 * split on the last dot: Tables/Forms/Reports use compound extensions (e.g. ".report.txt") that
 * contain a dot themselves, so a naive last-dot split would strip only ".txt". Shared by the
 * accessdb: URI parser and the on-disk AI mirror's reverse path -> object mapping.
 */
export function extFor(category: Category, fileName: string): string | undefined {
  return category === 'Modules'
    ? MODULE_EXTS.find((e) => fileName.endsWith(e))
    : (fileName.endsWith(EXT_FOR[category]) ? EXT_FOR[category] : undefined);
}

// Windows forbids more characters in real filenames than accessdb: URI paths do, has reserved
// device basenames, and strips trailing dots/spaces — encodeName (URI-only) is not enough here.
const FS_INVALID_CHARS = /[%<>:"/\\|?*\x00-\x1f]/g;
const FS_RESERVED_BASENAME = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i;

/** Escapes an Access object name into a real, safe Windows filename segment (no extension). */
export function encodeFsName(name: string): string {
  let s = name.replace(FS_INVALID_CHARS, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'));
  s = s.replace(/[ .]+$/g, (m) =>
    [...m].map((ch) => '%' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0')).join(''),
  );
  return FS_RESERVED_BASENAME.test(s) ? '_' + s : s;
}

export function decodeFsName(segment: string): string {
  const unprefixed = segment.startsWith('_') && FS_RESERVED_BASENAME.test(segment.slice(1)) ? segment.slice(1) : segment;
  return unprefixed.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

export function objectUri(key: string, category: Category, name: string, ext: string): vscode.Uri {
  return vscode.Uri.from({ scheme: SCHEME, path: `/${key}/${category}/${encodeName(name)}${ext}` });
}

export interface ParsedUri {
  key: string;
  category: Category;
  name: string;
  ext: string;
}

export function parseUri(uri: vscode.Uri): ParsedUri {
  const parts = uri.path.split('/').filter((p) => p.length > 0);
  if (parts.length !== 3 || !CATEGORIES.includes(parts[1] as Category)) {
    throw vscode.FileSystemError.FileNotFound(uri);
  }
  const category = parts[1] as Category;
  const fileName = parts[2];
  const ext = extFor(category, fileName);
  if (!ext) {
    throw vscode.FileSystemError.FileNotFound(uri);
  }
  return {
    key: parts[0],
    category,
    name: decodeName(fileName.slice(0, fileName.length - ext.length)),
    ext,
  };
}

/** Registry of open databases shared by the tree, filesystem provider and commands. */
export class DatabaseRegistry {
  private readonly dbs = new Map<string, OpenDatabase>();
  private readonly changeEmitter = new vscode.EventEmitter<void>();
  readonly onDidChange = this.changeEmitter.event;

  get all(): OpenDatabase[] {
    return [...this.dbs.values()];
  }

  get(key: string): OpenDatabase | undefined {
    return this.dbs.get(key);
  }

  getByPath(dbPath: string): OpenDatabase | undefined {
    return this.dbs.get(dbKeyFor(dbPath));
  }

  add(db: OpenDatabase): void {
    this.dbs.set(db.key, db);
    this.changeEmitter.fire();
  }

  remove(key: string): void {
    if (this.dbs.delete(key)) {
      this.changeEmitter.fire();
    }
  }

  notifyChanged(): void {
    this.changeEmitter.fire();
  }

  /** Human label for tab tooltips and tree: file name without directory. */
  labelFor(db: OpenDatabase): string {
    return path.basename(db.dbPath);
  }
}
