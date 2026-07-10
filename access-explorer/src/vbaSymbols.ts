import * as vscode from 'vscode';
import { Category, parseUri } from './model';

/**
 * Groups Sub/Function/Property procedures the way the VBA IDE's object/procedure dropdowns do:
 * a name containing an underscore ("cmdSave_Click") is a control's event handler — the part before
 * the underscore (including "Form"/"Report" for the object's own events, e.g. "Form_Load") becomes
 * the object bucket, the part after becomes the procedure label. Anything else falls under
 * "(General)". Works directly on the raw SaveAsText export for Forms/Reports too — the design
 * section's Begin/End property blocks never contain a line that matches a procedure declaration,
 * so no separate "find the code section" step is needed.
 */

const GENERAL = '(General)';

const DECL_RE =
  /^\s*(?:Private\s+|Public\s+|Friend\s+|Static\s+)*(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_]\w*)\s*\(/i;
const END_RE = /^\s*End\s+(Sub|Function|Property)\s*$/i;

interface Proc {
  name: string;
  kindWord: string;
  startLine: number;
  endLine: number;
}

function parseProcedures(text: string): Proc[] {
  const lines = text.split(/\r\n|\n/);
  const procs: Proc[] = [];
  let current: Proc | null = null;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (current) {
      if (END_RE.test(line)) {
        current.endLine = i;
        procs.push(current);
        current = null;
      }
      continue;
    }
    const m = DECL_RE.exec(line);
    if (m) {
      current = { name: m[2], kindWord: m[1].toLowerCase(), startLine: i, endLine: i };
    }
  }
  if (current) {
    current.endLine = lines.length - 1;
    procs.push(current);
  }
  return procs;
}

function splitName(name: string): { bucket: string; label: string } {
  const idx = name.indexOf('_');
  if (idx <= 0 || idx === name.length - 1) {
    return { bucket: GENERAL, label: name };
  }
  return { bucket: name.slice(0, idx), label: name.slice(idx + 1) };
}

function symbolKindFor(kindWord: string): vscode.SymbolKind {
  if (kindWord.startsWith('property')) {
    return vscode.SymbolKind.Property;
  }
  return kindWord === 'function' ? vscode.SymbolKind.Function : vscode.SymbolKind.Method;
}

export const CODELESS_CATEGORIES: ReadonlySet<Category> = new Set(['Tables', 'Queries', 'Macros']);

export interface ProcRange {
  /** "Click", or the bare name (e.g. a (General) proc) when the bucket has no sub-label. */
  label: string;
  range: vscode.Range;
  kindWord: string;
}

export interface ProcGroup {
  /** Control/object name (e.g. "cmdSave", "Form"), or the literal "(General)" bucket. */
  bucket: string;
  /** Covers every event of this bucket — from the first procedure's start to the last one's end. */
  range: vscode.Range;
  events: ProcRange[];
}

/**
 * Parses a document's raw text into the same control→events grouping used for the Outline
 * (see file header comment): one group per bucket, each with the range spanning all of its
 * events plus each event's own range. Shared by the DocumentSymbolProvider below and by the
 * "select code for AI" command, so there is exactly one place that understands this convention.
 */
export function groupProcedures(document: vscode.TextDocument): ProcGroup[] {
  const procs = parseProcedures(document.getText());
  if (procs.length === 0) {
    return [];
  }

  const buckets = new Map<string, Proc[]>();
  for (const proc of procs) {
    const { bucket } = splitName(proc.name);
    const list = buckets.get(bucket);
    if (list) {
      list.push(proc);
    } else {
      buckets.set(bucket, [proc]);
    }
  }

  const lineLength = (line: number): number => document.lineAt(line).text.length;
  const fullRange = (proc: Proc): vscode.Range =>
    new vscode.Range(proc.startLine, 0, proc.endLine, lineLength(proc.endLine));

  const bucketNames = [...buckets.keys()].sort((a, b) => {
    if (a === GENERAL) return -1;
    if (b === GENERAL) return 1;
    return a.localeCompare(b);
  });

  return bucketNames.map((bucket) => {
    const bucketProcs = buckets.get(bucket)!;
    const events = bucketProcs.map((proc) => {
      const { label } = splitName(proc.name);
      return {
        label: bucket === GENERAL ? proc.name : label,
        range: fullRange(proc),
        kindWord: proc.kindWord,
      };
    });
    const first = bucketProcs[0];
    const last = bucketProcs[bucketProcs.length - 1];
    return {
      bucket,
      range: new vscode.Range(first.startLine, 0, last.endLine, lineLength(last.endLine)),
      events,
    };
  });
}

export class VbaSymbolProvider implements vscode.DocumentSymbolProvider {
  provideDocumentSymbols(document: vscode.TextDocument): vscode.DocumentSymbol[] {
    let category: Category;
    try {
      ({ category } = parseUri(document.uri));
    } catch {
      return [];
    }
    if (CODELESS_CATEGORIES.has(category)) {
      return [];
    }

    return groupProcedures(document).map((group) => {
      const children = group.events.map(
        (event) =>
          new vscode.DocumentSymbol(
            event.label,
            '',
            symbolKindFor(event.kindWord),
            event.range,
            event.range,
          ),
      );
      // The bucket itself gets the FIRST child's selectionRange (now its full body), so picking
      // just the control (no specific event) already selects its first subroutine in full.
      const bucketSymbol = new vscode.DocumentSymbol(
        group.bucket,
        '',
        group.bucket === GENERAL ? vscode.SymbolKind.Module : vscode.SymbolKind.Class,
        group.range,
        children[0].selectionRange,
      );
      bucketSymbol.children = children;
      return bucketSymbol;
    });
  }
}
