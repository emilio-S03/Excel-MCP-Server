import { diagnoseConnectionLive } from './excel-live.js';

export async function diagnoseConnection(
  filePath?: string,
  _responseFormat: string = 'json'
): Promise<string> {
  const raw = await diagnoseConnectionLive(filePath);
  const steps = JSON.parse(raw);

  // Compute summary and healthy flag
  const stepKeys = Object.keys(steps).filter(k => k.startsWith('step'));
  const allPassed = stepKeys.every(k => steps[k].passed === true);
  const firstFailure = stepKeys.find(k => steps[k].passed === false);

  let summary: string;
  if (allPassed) {
    summary = 'All checks passed. Excel COM connection is healthy.';
  } else if (firstFailure) {
    summary = `Failed at ${firstFailure}: ${steps[firstFailure].message}`;
  } else {
    summary = 'Diagnostic completed with unknown status.';
  }

  const result = {
    healthy: allPassed,
    summary,
    steps,
  };

  return JSON.stringify(result, null, 2);
}
