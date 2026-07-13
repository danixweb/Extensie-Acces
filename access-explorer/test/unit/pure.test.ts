// pure.test.ts — headless unit tests for the vscode-free modules.
// Run via test/unit/run-unit.mjs (esbuild bundle, node:assert, TAP-like output).
//
// The form-design tests use a REAL SaveAsText export if present at
// test/unit/fixtures/form.def.txt (captured from a bridge suite run); otherwise
// they fall back to skipping, since a hand-built form export is not representative.

import * as assert from 'node:assert';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { splitModuleHeader, joinModuleHeader, synthesizeStandardHeader, splitFormHeader } from '../../src/vbaHeader';
import { extractSqlReferences, extractDomainReferences, extractCalledProcedureNames } from '../../src/mirrorDependencies';
import { parseDesignText } from '../../src/formDesignParser';
import { generateHtml } from '../../src/htmlMockupGenerator';

let n = 0;
let failed = 0;
function test(name: string, fn: () => void): void {
  n++;
  try {
    fn();
    console.log(`ok ${n} - ${name}`);
  } catch (err) {
    failed++;
    const msg = err instanceof Error ? err.message : String(err);
    console.log(`not ok ${n} - ${name}\n  # ${msg.replace(/\n/g, '\n  # ')}`);
  }
}
function skip(name: string, reason: string): void {
  n++;
  console.log(`ok ${n} - ${name} # SKIP ${reason}`);
}

// ---------- vbaHeader ----------

const STD_MODULE =
  'Attribute VB_Name = "Module1"\r\nOption Explicit\r\n\r\nSub Foo()\r\n10 Debug.Print "x"\r\nEnd Sub\r\n';

test('splitModuleHeader: modul standard — doar Attribute VB_Name în antet', () => {
  const { header, body } = splitModuleHeader(STD_MODULE);
  assert.strictEqual(header, 'Attribute VB_Name = "Module1"\r\n');
  assert.ok(body.startsWith('Option Explicit'));
  assert.strictEqual(joinModuleHeader(header, body), STD_MODULE);
});

const CLASS_MODULE =
  'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n' +
  'Attribute VB_Name = "clsFoo"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Exposed = False\r\n' +
  'Option Explicit\r\n\r\nPublic Sub Bar()\r\nEnd Sub\r\n';

test('splitModuleHeader: modul clasă — VERSION/BEGIN..END/Attribute în antet', () => {
  const { header, body } = splitModuleHeader(CLASS_MODULE);
  assert.ok(header.includes('VERSION 1.0 CLASS'));
  assert.ok(header.includes('END'));
  assert.ok(header.endsWith('Attribute VB_Exposed = False\r\n'));
  assert.ok(body.startsWith('Option Explicit'));
  assert.strictEqual(joinModuleHeader(header, body), CLASS_MODULE);
});

test('splitModuleHeader: Attribute din mijlocul corpului rămâne în corp', () => {
  const src =
    'Attribute VB_Name = "Module2"\r\nOption Explicit\r\n' +
    'Public Property Get Value() As Long\r\nAttribute Value.VB_UserMemId = 0\r\nEnd Property\r\n';
  const { header, body } = splitModuleHeader(src);
  assert.strictEqual(header, 'Attribute VB_Name = "Module2"\r\n');
  assert.ok(body.includes('Attribute Value.VB_UserMemId = 0'));
  assert.strictEqual(joinModuleHeader(header, body), src);
});

test('joinModuleHeader: normalizează LF din editor la CRLF', () => {
  assert.strictEqual(joinModuleHeader('H\r\n', 'a\nb\n'), 'H\r\na\r\nb\r\n');
});

test('synthesizeStandardHeader', () => {
  assert.strictEqual(synthesizeStandardHeader('Modul X'), 'Attribute VB_Name = "Modul X"\r\n');
});

const FORM_DEF =
  'Version =20\r\nVersionRequired =20\r\nBegin Form\r\n    Caption ="Test"\r\n' +
  '    Begin\r\n        Begin TextBox\r\n            Name ="txtA"\r\n        End\r\n    End\r\nEnd\r\n' +
  'CodeBehindForm\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Exposed = False\r\n' +
  'Option Explicit\r\n\r\nPrivate Sub txtA_Click()\r\nEnd Sub\r\n';

test('splitFormHeader: design + Attribute în antet, codul utilizator în corp', () => {
  const split = splitFormHeader(FORM_DEF);
  assert.ok(split, 'splitFormHeader a întors undefined');
  assert.ok(split!.header.includes('CodeBehindForm'));
  assert.ok(split!.header.endsWith('Attribute VB_Exposed = False\r\n'));
  assert.ok(split!.body.startsWith('Option Explicit'));
  assert.strictEqual(split!.header + split!.body, FORM_DEF);
});

test('splitFormHeader: fără CodeBehindForm → undefined (obiect read-only)', () => {
  assert.strictEqual(splitFormHeader('Version =20\r\nBegin Form\r\nEnd\r\n'), undefined);
});

// ---------- mirrorDependencies ----------

test('extractSqlReferences: FROM/JOIN cu paranteze drepte și virgule', () => {
  const refs = extractSqlReferences(
    'SELECT a.x, b.y FROM [Tabel Unu] AS a INNER JOIN Tabel2 AS b ON a.id=b.id, Tabel3 WHERE a.x>0;',
  );
  assert.deepStrictEqual(refs.sort(), ['Tabel Unu', 'Tabel2', 'Tabel3'].sort());
});

test('extractSqlReferences: INTO și UPDATE', () => {
  assert.ok(extractSqlReferences('SELECT * INTO [Arhiva 2024] FROM Sursa;').includes('Arhiva 2024'));
  assert.ok(extractSqlReferences('UPDATE Clienti SET x=1 WHERE id=2;').includes('Clienti'));
});

test('extractDomainReferences: DLookup/DCount cu domeniu literal', () => {
  const code =
    'v = DLookup("[col]", "dbo_Erori program", "id=1")\r\nn = DCount("*", "Clienti")\r\nm = DSum("x", variabila)';
  const refs = extractDomainReferences(code);
  assert.deepStrictEqual(refs.sort(), ['Clienti', 'dbo_Erori program'].sort());
});

test('extractCalledProcedureNames: găsește apelurile către proceduri cunoscute', () => {
  const known = new Set(['SCRIEEROARE', 'ALTAPROC']);
  const refs = extractCalledProcedureNames('If x Then Call ScrieEroare(Err.Number, "m")', known);
  assert.deepStrictEqual(refs, ['SCRIEEROARE']);
});

// ---------- formDesignParser + htmlMockupGenerator (fixture real, dacă există) ----------

const fixturePath = path.join(process.env.AX_FIXTURES ?? path.join(__dirname, 'fixtures'), 'form.def.txt');
if (fs.existsSync(fixturePath)) {
  const raw = fs.readFileSync(fixturePath, 'utf8');
  test('parseDesignText: parsează un export real de formular', () => {
    const design = parseDesignText(raw);
    assert.ok(design, 'parseDesignText a întors undefined pe un export real');
    assert.ok(design!.sections.length > 0, 'niciun section parsat');
    const controls = design!.sections.flatMap((s) => s.controls);
    assert.ok(controls.length > 0, 'niciun control parsat');
  });
  test('generateHtml: mockup-ul conține conținutul randat al controalelor', () => {
    const design = parseDesignText(raw)!;
    const html = generateHtml(design, 'FixtureForm');
    assert.ok(html.length > 500, 'HTML suspect de scurt');
    const controls = design.sections.flatMap((s) => s.controls);
    // TextBox/ComboBox/ListBox se randează cu [controlSource], nu cu numele;
    // Label/CommandButton se randează cu caption (sau nume). Verificăm ce chiar se afișează.
    const captioned = controls.find((c) => (c.type === 'Label' || c.type === 'CommandButton') && (c.caption ?? c.name));
    const bound = controls.find(
      (c) => (c.type === 'TextBox' || c.type === 'ComboBox' || c.type === 'ListBox') && c.controlSource,
    );
    assert.ok(captioned || bound, 'fixture fără controale randabile cu text');
    if (captioned) assert.ok(html.includes(captioned.caption ?? captioned.name), `caption-ul '${captioned.caption ?? captioned.name}' lipsește din HTML`);
    if (bound) assert.ok(html.includes(`[${bound.controlSource}]`), `sursa '[${bound.controlSource}]' lipsește din HTML`);
    assert.ok(html.includes(`class="section"`), 'secțiunile lipsesc din HTML');
  });
} else {
  skip('parseDesignText: export real de formular', 'fixture lipsă (test/unit/fixtures/form.def.txt)');
  skip('generateHtml: mockup', 'fixture lipsă');
}

console.log(`1..${n}`);
console.log(`# unit: ${n - failed} pass, ${failed} fail`);
if (failed > 0) process.exitCode = 1;
