/**
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

// Bundles index.ts with Univer and React 18 into ../runtime-dist/ (see
// index.ts for why). Superset's build, build-dev, dev and dev-server scripts
// run this first through `npm run build-univer-runtime`.
import { execSync } from 'child_process';
import { existsSync, rmSync, writeFileSync } from 'fs';
import { dirname, resolve } from 'path';
import { fileURLToPath } from 'url';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(here, '../runtime-dist');

if (!existsSync(resolve(here, 'node_modules'))) {
  execSync('npm ci --no-audit --no-fund', { cwd: here, stdio: 'inherit' });
}

rmSync(outDir, { recursive: true, force: true });

// imported after the install above, esbuild comes from this project
const { build } = await import('esbuild');

await build({
  absWorkingDir: here,
  entryPoints: { 'univer-runtime': 'index.ts' },
  outdir: outDir,
  // keeps Univer's lazily loaded parts (like hyphenation dictionaries) out of
  // the main file; webpack turns them into chunks loaded on demand
  splitting: true,
  chunkNames: 'chunks/[name]-[hash]',
  bundle: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2020',
  minify: true,
  legalComments: 'external',
  define: { 'process.env.NODE_ENV': '"production"' },
  logLevel: 'info',
});

writeFileSync(
  resolve(outDir, 'univer-runtime.d.ts'),
  "import type { MountUniver } from '../src/runtimeTypes';\n\nexport declare const mountUniver: MountUniver;\n",
);
