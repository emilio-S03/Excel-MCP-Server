import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(__dirname, '..');

test('manifest.json matches actual tool registrations', () => {
  // Will throw with non-zero exit if drift exists.
  execFileSync('node', ['scripts/generate-manifest.mjs', '--check'], {
    cwd: repoRoot,
    stdio: 'pipe',
  });
});

test('manifest declares the v3 user_config schema', () => {
  const manifest = JSON.parse(readFileSync(join(repoRoot, 'manifest.json'), 'utf8'));
  assert.equal(manifest.version, '3.3.0');
  assert.ok(manifest.user_config.allowedDirectories, 'should declare allowedDirectories');
  assert.equal(manifest.user_config.allowedDirectories.type, 'directory');
  assert.equal(manifest.user_config.allowedDirectories.multiple, true);
  assert.equal(
    manifest.server.mcp_config.env.EXCEL_ALLOWED_DIRS,
    '${user_config.allowedDirectories}',
    'should template-substitute allowedDirectories into EXCEL_ALLOWED_DIRS'
  );
});

test('every registered tool has a name + description in manifest', () => {
  const manifest = JSON.parse(readFileSync(join(repoRoot, 'manifest.json'), 'utf8'));
  for (const t of manifest.tools) {
    assert.ok(t.name, 'tool must have name');
    assert.ok(t.description, `tool ${t.name} must have description`);
    assert.ok(t.name.startsWith('excel_'), `tool name should start with excel_ (got ${t.name})`);
  }
});
