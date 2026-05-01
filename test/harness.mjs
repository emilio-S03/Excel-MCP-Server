/**
 * Reusable MCP wire-protocol test harness.
 * Spawns the built server over stdio and provides initialize / call helpers.
 */
import { spawn } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
export const repoRoot = join(__dirname, '..');
export const serverPath = join(repoRoot, 'dist', 'index.js');

export async function startServer(envOverrides = {}) {
  const responses = new Map();
  let stdoutBuf = '';
  let stderrBuf = '';
  let nextId = 1;
  let settled = false;

  const child = spawn('node', [serverPath], {
    stdio: ['pipe', 'pipe', 'pipe'],
    env: { ...process.env, ...envOverrides },
  });

  child.stdout.on('data', (chunk) => {
    stdoutBuf += chunk.toString();
    let nl;
    while ((nl = stdoutBuf.indexOf('\n')) !== -1) {
      const line = stdoutBuf.slice(0, nl).trim();
      stdoutBuf = stdoutBuf.slice(nl + 1);
      if (!line) continue;
      try {
        const msg = JSON.parse(line);
        if (msg.id !== undefined) responses.set(msg.id, msg);
      } catch {
        // ignore non-JSON
      }
    }
  });

  child.stderr.on('data', (c) => {
    stderrBuf += c.toString();
  });

  function send(method, params, withId = true) {
    const obj = withId
      ? { jsonrpc: '2.0', id: nextId++, method, params }
      : { jsonrpc: '2.0', method, params };
    child.stdin.write(JSON.stringify(obj) + '\n');
    return obj.id;
  }

  async function awaitResponse(id, timeoutMs = 10000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (responses.has(id)) return responses.get(id);
      await new Promise((r) => setTimeout(r, 25));
    }
    throw new Error(`timeout waiting for response ${id}\n--- server stderr ---\n${stderrBuf}`);
  }

  // Handshake
  const initId = send('initialize', {
    protocolVersion: '2024-11-05',
    capabilities: {},
    clientInfo: { name: 'mcp-test-harness', version: '0.0.1' },
  });
  await awaitResponse(initId);
  send('notifications/initialized', {}, false);

  return {
    callTool: async (name, args, timeoutMs = 15000) => {
      const id = send('tools/call', { name, arguments: args });
      const resp = await awaitResponse(id, timeoutMs);
      return resp;
    },
    listTools: async () => {
      const id = send('tools/list', {});
      const resp = await awaitResponse(id);
      return resp.result.tools;
    },
    stop: () => {
      if (settled) return;
      settled = true;
      child.kill();
    },
    getStderr: () => stderrBuf,
  };
}
