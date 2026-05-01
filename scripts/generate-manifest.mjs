#!/usr/bin/env node
/**
 * Regenerate manifest.json `tools[]` from the server's actual tools/list response.
 *
 * Spawns dist/index.js over stdio, performs a real MCP handshake + tools/list,
 * and rewrites the tools array. This makes drift between manifest and code
 * impossible: the manifest is *derived* from the running server.
 *
 * Usage:
 *   node scripts/generate-manifest.mjs           # rewrite manifest.json in place
 *   node scripts/generate-manifest.mjs --check   # exit 1 if manifest would change (CI gate)
 *   node scripts/generate-manifest.mjs --print   # print to stdout, don't write
 */
import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(__dirname, '..');
const manifestPath = join(repoRoot, 'manifest.json');
const serverPath = join(repoRoot, 'dist', 'index.js');

const args = new Set(process.argv.slice(2));
const CHECK = args.has('--check');
const PRINT = args.has('--print');

async function listToolsFromServer() {
  return new Promise((resolve, reject) => {
    const child = spawn('node', [serverPath], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    let stdoutBuffer = '';
    let stderrBuffer = '';
    const responses = new Map();
    let settled = false;

    child.stdout.on('data', (chunk) => {
      stdoutBuffer += chunk.toString();
      let nlIdx;
      while ((nlIdx = stdoutBuffer.indexOf('\n')) !== -1) {
        const line = stdoutBuffer.slice(0, nlIdx).trim();
        stdoutBuffer = stdoutBuffer.slice(nlIdx + 1);
        if (!line) continue;
        try {
          const msg = JSON.parse(line);
          if (msg.id !== undefined) responses.set(msg.id, msg);
        } catch {
          // server may emit non-JSON log lines on stdout; ignore
        }
      }
    });

    child.stderr.on('data', (chunk) => {
      stderrBuffer += chunk.toString();
    });

    child.on('error', (err) => {
      if (!settled) {
        settled = true;
        reject(err);
      }
    });

    child.on('exit', (code) => {
      if (!settled) {
        settled = true;
        reject(new Error(`server exited early (code ${code}). stderr:\n${stderrBuffer}`));
      }
    });

    function send(obj) {
      child.stdin.write(JSON.stringify(obj) + '\n');
    }

    async function awaitResponse(id, timeoutMs = 10000) {
      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        if (responses.has(id)) return responses.get(id);
        await new Promise((r) => setTimeout(r, 25));
      }
      throw new Error(`timeout waiting for response ${id}`);
    }

    (async () => {
      try {
        send({
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
          params: {
            protocolVersion: '2024-11-05',
            capabilities: {},
            clientInfo: { name: 'manifest-generator', version: '1.0.0' },
          },
        });
        await awaitResponse(1);

        send({ jsonrpc: '2.0', method: 'notifications/initialized' });

        send({ jsonrpc: '2.0', id: 2, method: 'tools/list' });
        const listResp = await awaitResponse(2);

        if (listResp.error) {
          throw new Error(`tools/list returned error: ${JSON.stringify(listResp.error)}`);
        }

        settled = true;
        child.kill();
        resolve(listResp.result.tools);
      } catch (err) {
        if (!settled) {
          settled = true;
          child.kill();
          reject(err);
        }
      }
    })();
  });
}

function deriveManifestTools(tools) {
  return tools.map((t) => ({
    name: t.name,
    description: t.description,
  }));
}

async function main() {
  const tools = await listToolsFromServer();
  const newToolsBlock = deriveManifestTools(tools);

  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
  const oldToolsBlock = manifest.tools ?? [];

  manifest.tools = newToolsBlock;

  const serialized = JSON.stringify(manifest, null, 2) + '\n';

  if (PRINT) {
    process.stdout.write(serialized);
    return;
  }

  if (CHECK) {
    const current = readFileSync(manifestPath, 'utf8');
    if (current === serialized) {
      console.log(`OK: manifest.json is in sync (${newToolsBlock.length} tools)`);
      return;
    }
    const oldNames = new Set(oldToolsBlock.map((t) => t.name));
    const newNames = new Set(newToolsBlock.map((t) => t.name));
    const added = [...newNames].filter((n) => !oldNames.has(n));
    const removed = [...oldNames].filter((n) => !newNames.has(n));
    console.error(`DRIFT: manifest.json out of sync.`);
    console.error(`  manifest had ${oldToolsBlock.length} tools, server exposes ${newToolsBlock.length}`);
    if (added.length) console.error(`  + added (${added.length}): ${added.join(', ')}`);
    if (removed.length) console.error(`  - removed (${removed.length}): ${removed.join(', ')}`);
    console.error(`Run: node scripts/generate-manifest.mjs`);
    process.exit(1);
  }

  writeFileSync(manifestPath, serialized);
  const oldCount = oldToolsBlock.length;
  const newCount = newToolsBlock.length;
  if (oldCount === newCount) {
    console.log(`Wrote manifest.json: ${newCount} tools (no count change; descriptions may differ)`);
  } else {
    console.log(`Wrote manifest.json: ${oldCount} -> ${newCount} tools`);
  }
}

main().catch((err) => {
  console.error('manifest generation failed:', err);
  process.exit(2);
});
