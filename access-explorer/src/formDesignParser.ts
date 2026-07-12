/**
 * Parser for the Form/Report "design section" produced by Access's SaveAsText — the
 * `Begin Form|Report ... End` block of nested property groups that `vbaHeader.ts` already
 * isolates (verbatim, unparsed) as `entry.header`. Grammar reverse-engineered from real
 * exports (no official spec exists); see calibration notes below.
 *
 * Structural facts confirmed against real exports:
 * - A bare `Begin` (no keyword) is a transparent grouping node: it appears once directly
 *   under the root to hold every section/control, and once under each bound control to hold
 *   its attached Label. It contributes no data of its own — recurse through it.
 * - A property can itself be a nested `Begin`/`End` block (`GUID = Begin ... End`,
 *   `NameMap = Begin ...`) holding opaque hex data (GUIDs, print settings). None of that is
 *   needed for a visual mockup — skip to the matching `End` and discard the value.
 * - Every control TYPE used anywhere in the object gets one `Begin <Type> ... End` "default
 *   style" template as a direct child of the root's group, holding only its default
 *   properties — never a `Name`. Real placed elements (sections and controls) always carry
 *   a `Name` property. That single fact reliably tells a real element apart from a template,
 *   without needing to know every control keyword in advance.
 * - Sections (FormHeader/PageHeader/GroupHeader.../Detail/GroupFooter.../PageFooter/
 *   FormFooter — Report headers/footers are internally still typed `FormHeader`/`FormFooter`,
 *   distinguished only by `Name="ReportHeader"`/`"ReportFooter"`) are exactly the named nodes
 *   one level below the root group, always emitted in on-screen top-to-bottom order — so file
 *   order can be used directly as display order, no name-based ordering table needed.
 * - Controls are every other named node, at any depth (an attached Label sits one level below
 *   its owning control) — collected as siblings within their enclosing section.
 */

export interface DesignControl {
  /** Raw `Begin <Type>` keyword (TextBox, Label, CommandButton, Subform, ...). */
  type: string;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  caption?: string;
  controlSource?: string;
  /** Subform/Subreport only: e.g. "Form.2 cautare clisee subform1". */
  sourceObject?: string;
  foreColor?: string;
  backColor?: string;
  fontName?: string;
  fontSize?: number;
  fontWeight?: number;
  borderStyle: 'none' | 'solid';
}

export interface DesignSection {
  name: string;
  height: number;
  backColor?: string;
  controls: DesignControl[];
}

export interface ParsedDesign {
  objectType: 'Form' | 'Report';
  width: number;
  caption?: string;
  recordSource?: string;
  sections: DesignSection[];
}

interface RawNode {
  /** null = anonymous `Begin` grouping node (transparent). */
  keyword: string | null;
  props: Map<string, string>;
  children: RawNode[];
}

const ROOT_LINE = /^\s*Begin\s+(Form|Report)\s*$/;
const BEGIN_TYPED = /^\s*Begin\s+(\S+)\s*$/;
const BEGIN_BARE = /^\s*Begin\s*$/;
const END_LINE = /^\s*End\s*$/;
/** A property whose value is itself an opaque nested block, e.g. `GUID = Begin`. */
const PROP_OPAQUE_BLOCK = /^\s*(\w+)\s*=\s*Begin\s*$/;
const PROP_LINE = /^\s*(\w+)\s*=\s*(.*)$/;

function unquote(value: string): string {
  if (value.length >= 2 && value.startsWith('"') && value.endsWith('"')) {
    return value.slice(1, -1).replace(/""/g, '"');
  }
  return value;
}

function parseBlock(lines: string[], cursor: { i: number }, keyword: string | null): RawNode {
  const node: RawNode = { keyword, props: new Map(), children: [] };
  while (cursor.i < lines.length) {
    const line = lines[cursor.i].trim();
    if (line === '') {
      cursor.i++;
      continue;
    }
    if (END_LINE.test(line)) {
      cursor.i++;
      return node;
    }
    const typed = BEGIN_TYPED.exec(line);
    if (typed) {
      cursor.i++;
      node.children.push(parseBlock(lines, cursor, typed[1]));
      continue;
    }
    if (BEGIN_BARE.test(line)) {
      cursor.i++;
      node.children.push(parseBlock(lines, cursor, null));
      continue;
    }
    if (PROP_OPAQUE_BLOCK.test(line)) {
      cursor.i++;
      // Opaque hex/GUID payload, not needed for a visual mockup — skip to its `End`.
      while (cursor.i < lines.length && !END_LINE.test(lines[cursor.i].trim())) {
        cursor.i++;
      }
      cursor.i++;
      continue;
    }
    const prop = PROP_LINE.exec(line);
    if (prop) {
      node.props.set(prop[1], unquote(prop[2].trim()));
      cursor.i++;
      continue;
    }
    // Unrecognized line (stray continuation, comment-like content) — ignore, keep going.
    cursor.i++;
  }
  return node;
}

function num(props: Map<string, string>, key: string, fallback: number): number {
  const v = props.get(key);
  if (v === undefined) {
    return fallback;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

/** 1440 twips/inch at 96 DPI = 15 twips/px. */
function twipsToPx(twips: number): number {
  return twips / 15;
}

/**
 * Access stores an OLE color as a signed 32-bit BGR integer. Negative values carry the
 * 0x80000000 "system color" flag (index into the Windows theme, e.g. button face/window
 * text) — resolving those exactly would need the live Windows theme, which isn't available
 * here, so they're left undefined and the generated CSS falls back to a neutral default.
 */
function colorCss(props: Map<string, string>, key: string): string | undefined {
  const v = props.get(key);
  if (v === undefined) {
    return undefined;
  }
  const n = Number(v);
  if (!Number.isFinite(n) || n < 0) {
    return undefined;
  }
  const r = n & 0xff;
  const g = (n >> 8) & 0xff;
  const b = (n >> 16) & 0xff;
  return `rgb(${r}, ${g}, ${b})`;
}

function toControl(node: RawNode): DesignControl {
  const p = node.props;
  return {
    type: node.keyword ?? 'Unknown',
    name: p.get('Name') ?? '',
    left: twipsToPx(num(p, 'Left', 0)),
    top: twipsToPx(num(p, 'Top', 0)),
    width: twipsToPx(num(p, 'Width', 1000)),
    height: twipsToPx(num(p, 'Height', 300)),
    caption: p.get('Caption'),
    controlSource: p.get('ControlSource'),
    sourceObject: p.get('SourceObject'),
    foreColor: colorCss(p, 'ForeColor'),
    backColor: colorCss(p, 'BackColor'),
    fontName: p.get('FontName'),
    fontSize: p.has('FontSize') ? num(p, 'FontSize', 0) : undefined,
    fontWeight: p.has('FontWeight') ? num(p, 'FontWeight', 0) : undefined,
    borderStyle: num(p, 'BorderLineStyle', 0) === 0 ? 'none' : 'solid',
  };
}

/**
 * Walks the tree collecting real (named) sections and controls, transparently flattening
 * every anonymous `Begin` group and discarding every unnamed "default style" template.
 */
function collect(node: RawNode, currentSection: DesignSection | null, sections: DesignSection[]): void {
  for (const child of node.children) {
    if (child.keyword === null) {
      collect(child, currentSection, sections);
      continue;
    }
    if (!child.props.has('Name')) {
      // Unnamed control-type default template — not a placed element.
      continue;
    }
    if (currentSection === null) {
      const section: DesignSection = {
        name: child.props.get('Name')!,
        height: twipsToPx(num(child.props, 'Height', 0)),
        backColor: colorCss(child.props, 'BackColor'),
        controls: [],
      };
      sections.push(section);
      collect(child, section, sections);
    } else {
      currentSection.controls.push(toControl(child));
      // Recurse (unchanged section) to also pick up an attached Label, if any.
      collect(child, currentSection, sections);
    }
  }
}

/** Returns undefined if `raw` doesn't contain a recognizable `Begin Form|Report` design section. */
export function parseDesignText(raw: string): ParsedDesign | undefined {
  const lines = raw.split(/\r\n|\n/);
  const rootIdx = lines.findIndex((l) => ROOT_LINE.test(l));
  if (rootIdx === -1) {
    return undefined;
  }
  const objectType = ROOT_LINE.exec(lines[rootIdx])![1] as 'Form' | 'Report';
  const cursor = { i: rootIdx + 1 };
  const root = parseBlock(lines, cursor, objectType);
  const sections: DesignSection[] = [];
  collect(root, null, sections);
  return {
    objectType,
    width: twipsToPx(num(root.props, 'Width', 0)),
    caption: root.props.get('Caption'),
    recordSource: root.props.get('RecordSource'),
    sections,
  };
}
