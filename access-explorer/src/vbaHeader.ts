/**
 * SaveAsText exports carry a leading header block that must be preserved verbatim
 * when writing back with LoadFromText, but should be hidden from the user:
 *
 *   VERSION 1.0 CLASS            (class modules only)
 *   BEGIN
 *     MultiUse = -1  'True
 *   END
 *   Attribute VB_Name = "clsFoo"
 *   Attribute VB_GlobalNameSpace = False
 *   ...
 *   Option Explicit              <-- user code starts here
 *
 * Standard modules have only leading `Attribute VB_Name = ...` lines. Only the
 * LEADING block is consumed — mid-file lines such as
 * `Attribute Value.VB_UserMemId = 0` belong to the body and stay untouched.
 */

const HEADER_LINE = /^(VERSION \d+\.\d+( CLASS)?\s*$|BEGIN\b|\s+\w+ = |END\s*$|Attribute VB_\w+ = )/;

export interface SplitModule {
  header: string;
  body: string;
}

export function splitModuleHeader(exported: string): SplitModule {
  const lines = exported.split(/\r\n|\n/);
  let i = 0;
  let inBeginBlock = false;
  while (i < lines.length) {
    const line = lines[i];
    if (inBeginBlock) {
      i++;
      if (/^END\s*$/.test(line)) {
        inBeginBlock = false;
      }
      continue;
    }
    if (/^BEGIN\b/.test(line)) {
      inBeginBlock = true;
      i++;
      continue;
    }
    if (HEADER_LINE.test(line) && /^(VERSION|Attribute VB_)/.test(line)) {
      i++;
      continue;
    }
    break;
  }
  return {
    header: lines.slice(0, i).join('\r\n') + (i > 0 ? '\r\n' : ''),
    body: lines.slice(i).join('\r\n'),
  };
}

/** Reassembles a full SaveAsText payload from a cached header and the edited body. */
export function joinModuleHeader(header: string, body: string): string {
  // Normalize editor line endings back to CRLF — VBA is strictly CRLF.
  const crlfBody = body.replace(/\r\n|\n/g, '\r\n');
  return header + crlfBody;
}

/** Fallback when no header is cached: synthesize a minimal standard-module header. */
export function synthesizeStandardHeader(moduleName: string): string {
  return `Attribute VB_Name = "${moduleName}"\r\n`;
}
