import { afterAll, beforeAll } from 'vitest';

/**
 * Opts a test file into pinned, byte-identical output, the way a caller would with
 * --reproducible. Output is real-time by default, so tests that compare bytes, hashes
 * or fixed ids across runs must ask for it.
 */
export function useReproducibleOutput(): void {
  let savedOn: string | undefined;
  let savedOff: string | undefined;
  beforeAll(() => {
    savedOn = process.env.ASTRA_REPRODUCIBLE;
    savedOff = process.env.ASTRA_NO_REPRODUCIBLE;
    delete process.env.ASTRA_NO_REPRODUCIBLE;
    process.env.ASTRA_REPRODUCIBLE = '1';
  });
  afterAll(() => {
    if (savedOn === undefined) delete process.env.ASTRA_REPRODUCIBLE;
    else process.env.ASTRA_REPRODUCIBLE = savedOn;
    if (savedOff === undefined) delete process.env.ASTRA_NO_REPRODUCIBLE;
    else process.env.ASTRA_NO_REPRODUCIBLE = savedOff;
  });
}
