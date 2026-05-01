/**
 * Centralized platform-aware error messages for tools that don't yet have
 * full cross-platform parity. Each error tells the user *what to do next*
 * instead of just "Requires Windows".
 *
 * Categories:
 *   - WIN_ONLY_USE_OFFICE_SCRIPT: Mac users should use Microsoft Office Scripts (cloud) instead.
 *   - WIN_ONLY_FILE_MODE_ALT: There's an existing file-mode tool that does ~80% of this on Mac.
 *   - WIN_ONLY_NO_ALT: Genuinely Windows-only with no Mac path. Document this.
 *   - PENDING_MAC_PORT: Will land in a future Mac-parity batch.
 */
import { platform } from 'os';

const IS_MAC = platform() === 'darwin';
const IS_LINUX = platform() === 'linux';

function makeError(toolName: string, body: string): Error {
  const platformLabel = IS_MAC ? 'macOS' : IS_LINUX ? 'Linux' : 'this platform';
  const err = new Error(
    `${toolName}: not yet supported on ${platformLabel}. ${body} ` +
    `Run \`excel_check_environment\` for a full list of what works on this machine.`
  );
  (err as any).code = 'PLATFORM_UNSUPPORTED';
  (err as any).platform = process.platform;
  return err;
}

export function winOnlyUseOfficeScript(toolName: string, capability: string): Error {
  return makeError(
    toolName,
    `${capability} via the desktop API is Windows-only — Microsoft disabled the equivalent path on Mac years ago. ` +
    `On macOS, use Microsoft Office Scripts (web Excel only — https://learn.microsoft.com/office/dev/scripts/) ` +
    `which works in Excel for the Web and via the Power Automate connector.`
  );
}

export function winOnlyFileModeAlt(toolName: string, capability: string, fileModeTool: string): Error {
  return makeError(
    toolName,
    `Live ${capability} requires Excel to be running with the file open, which currently only works on Windows. ` +
    `On macOS, close the file in Excel and use the file-mode tool \`${fileModeTool}\` instead — it edits the .xlsx directly and the changes appear when you reopen the file.`
  );
}

export function winOnlyNoAlt(toolName: string, capability: string): Error {
  return makeError(
    toolName,
    `${capability} relies on the Windows COM interface and has no AppleScript or JXA equivalent on macOS Excel.`
  );
}

export function pendingMacPort(toolName: string, etaBatch: string): Error {
  return makeError(
    toolName,
    `An AppleScript implementation is planned for ${etaBatch}. Track parity status in docs/PLATFORM_PARITY.md.`
  );
}
