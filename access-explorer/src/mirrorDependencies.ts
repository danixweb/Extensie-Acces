/**
 * Regex-based (not a real SQL/VBA parser) extraction of an object's dependencies, used to expand
 * on-demand mirroring beyond the single clicked object: a query's SQL references other
 * tables/queries, and VBA code references tables/queries via domain-aggregate functions and other
 * procedures via calls. False negatives (dynamically-built SQL/domain strings, calls resolved
 * through late binding) and rare false positives (a string literal that happens to look like a
 * name) are an accepted trade-off — a full parser would be disproportionate to the goal, which is
 * "mirror the objects actually related to what was opened," not exhaustive static analysis.
 */

const SQL_TABLE_KEYWORD_RE = /\bFROM\s+([\s\S]*?)(?:\bWHERE\b|\bGROUP\s+BY\b|\bORDER\s+BY\b|\bHAVING\b|;|$)/gi;
const SQL_JOIN_SPLIT_RE = /,|\bINNER\s+JOIN\b|\bLEFT\s+JOIN\b|\bRIGHT\s+JOIN\b|\bJOIN\b/gi;
const SQL_LEADING_NAME_RE = /^(\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_]*)/;
const SQL_INTO_RE = /\bINTO\s+(\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_]*)/gi;
const SQL_UPDATE_RE = /\bUPDATE\s+(\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_]*)/gi;

function unbracket(raw: string): string {
  return raw.startsWith('[') && raw.endsWith(']') ? raw.slice(1, -1) : raw;
}

/** Tables/queries referenced by a query's SQL: FROM/JOIN (including comma-joins), INTO, UPDATE. */
export function extractSqlReferences(sql: string): string[] {
  const found = new Set<string>();

  let m: RegExpExecArray | null;
  const fromRe = new RegExp(SQL_TABLE_KEYWORD_RE);
  while ((m = fromRe.exec(sql))) {
    const clause = m[1];
    for (const part of clause.split(SQL_JOIN_SPLIT_RE)) {
      const tok = SQL_LEADING_NAME_RE.exec(part.trim());
      if (tok) {
        found.add(unbracket(tok[1]));
      }
    }
  }
  const intoRe = new RegExp(SQL_INTO_RE);
  while ((m = intoRe.exec(sql))) {
    found.add(unbracket(m[1]));
  }
  const updateRe = new RegExp(SQL_UPDATE_RE);
  while ((m = updateRe.exec(sql))) {
    found.add(unbracket(m[1]));
  }
  return [...found];
}

const DOMAIN_FN_RE =
  /\bD(?:Count|Sum|Lookup|Avg|Min|Max|First|Last)\s*\(\s*(?:"(?:[^"]|"")*"|[^,()]*)\s*,\s*"((?:[^"]|"")*)"/gi;

/** Tables/queries named as the domain (2nd argument) of a D-function: DCount/DSum/DLookup/etc.
 *  Only resolves a literal quoted string — a domain built from a variable can't be resolved statically. */
export function extractDomainReferences(code: string): string[] {
  const found = new Set<string>();
  const re = new RegExp(DOMAIN_FN_RE);
  let m: RegExpExecArray | null;
  while ((m = re.exec(code))) {
    const name = m[1].replace(/""/g, '"');
    if (name) {
      found.add(name);
    }
  }
  return [...found];
}

const IDENTIFIER_RE = /\b[A-Za-z_][A-Za-z0-9_]*\b/g;

/** Uppercased names, among `knownProcedureNames` (already uppercased), that appear anywhere in
 *  `code` — used to find which other modules' procedures a routine calls. Matches on plain
 *  identifier presence, not real call-site syntax, so a local variable that happens to share a
 *  name with an unrelated procedure elsewhere is an accepted (harmless) false positive. */
export function extractCalledProcedureNames(code: string, knownProcedureNames: ReadonlySet<string>): string[] {
  if (knownProcedureNames.size === 0) {
    return [];
  }
  const found = new Set<string>();
  const re = new RegExp(IDENTIFIER_RE);
  let m: RegExpExecArray | null;
  while ((m = re.exec(code))) {
    const upper = m[0].toUpperCase();
    if (knownProcedureNames.has(upper)) {
      found.add(upper);
    }
  }
  return [...found];
}
