// run-unit.mjs — bundles and runs the headless unit tests for the vscode-free
// TS modules (vbaHeader, mirrorDependencies, formDesignParser, htmlMockupGenerator).
//   node test/unit/run-unit.mjs
import { buildSync } from 'esbuild';
import * as os from 'node:os';
import * as path from 'node:path';
import * as fs from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(os.tmpdir(), 'access-explorer-tests');
fs.mkdirSync(outDir, { recursive: true });
const outfile = path.join(outDir, `unit-${Date.now()}.cjs`);

buildSync({
  entryPoints: [path.join(__dirname, 'pure.test.ts')],
  bundle: true,
  platform: 'node',
  format: 'cjs',
  outfile,
  logLevel: 'silent',
});

// The bundle runs from %TEMP%, so __dirname there is useless for locating fixtures.
process.env.AX_FIXTURES = path.join(__dirname, 'fixtures');
await import(pathToFileURL(outfile));
