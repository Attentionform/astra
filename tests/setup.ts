// Global test setup for Vitest
// Ensures in-memory localStorage stub exists in Node.js test environment

let store: Record<string, string> = {};
const storageStub = {
  getItem: (k: string) => store[k] ?? null,
  setItem: (k: string, v: string) => { store[k] = String(v); },
  removeItem: (k: string) => { delete store[k]; },
  clear: () => { store = {}; },
  get length() { return Object.keys(store).length; },
  key: (i: number) => Object.keys(store)[i] ?? null,
};

if (!globalThis.localStorage) {
  (globalThis as any).localStorage = storageStub;
}
if (!(global as any).localStorage) {
  (global as any).localStorage = storageStub;
}
