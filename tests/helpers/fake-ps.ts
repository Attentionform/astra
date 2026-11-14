import * as path from 'path';
import { fileURLToPath } from 'url';

const here = path.dirname(fileURLToPath(import.meta.url));
export const FAKE_PS = path.join(here, 'fake-powershell.mjs');
/** Options that make the engine start the fake instead of powershell.exe. */
export const fakePsHooks = { powershellBin: process.execPath, powershellArgs: [FAKE_PS] };

/** Set FAKE_PS_* variables for one test and restore them afterwards. */
export function withFakePsEnv<T>(env: Record<string, string | undefined>, fn: () => Promise<T>): Promise<T> {
  const saved: Record<string, string | undefined> = {};
  for (const k of Object.keys(env)) {
    saved[k] = process.env[k];
    if (env[k] === undefined) delete process.env[k];
    else process.env[k] = env[k];
  }
  return fn().finally(() => {
    for (const k of Object.keys(saved)) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
  });
}
