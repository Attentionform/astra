// Small, dependency-free statistics helpers. Deterministic bootstrap (seeded).
export function quantile(sorted, q) {
  if (!sorted.length) return NaN;
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}
export function summarize(values) {
  const v = values.filter((x) => Number.isFinite(x)).slice().sort((a, b) => a - b);
  const n = v.length;
  if (!n) return { n: 0 };
  const mean = v.reduce((a, b) => a + b, 0) / n;
  const sd = n > 1 ? Math.sqrt(v.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1)) : 0;
  return {
    n, min: v[0], p25: quantile(v, 0.25), median: quantile(v, 0.5), p75: quantile(v, 0.75),
    p95: quantile(v, 0.95), max: v[n - 1], mean, cv: mean ? sd / mean : 0,
  };
}
function mulberry32(a) {
  return () => { a |= 0; a = (a + 0x6d2b79f5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}
export function bootstrapMedianCI(values, { resamples = 10000, seed = 20261111, alpha = 0.05 } = {}) {
  const v = values.filter((x) => Number.isFinite(x));
  if (v.length < 2) return [NaN, NaN];
  const rnd = mulberry32(seed);
  const meds = new Array(resamples);
  for (let r = 0; r < resamples; r++) {
    const s = new Array(v.length);
    for (let i = 0; i < v.length; i++) s[i] = v[Math.floor(rnd() * v.length)];
    s.sort((a, b) => a - b);
    meds[r] = quantile(s, 0.5);
  }
  meds.sort((a, b) => a - b);
  return [quantile(meds, alpha / 2), quantile(meds, 1 - alpha / 2)];
}
// Wilson score interval for a proportion (95% by default).
export function wilson(k, n, z = 1.959964) {
  if (n === 0) return { k, n, rate: NaN, lo: NaN, hi: NaN };
  const p = k / n, d = 1 + (z * z) / n;
  const c = (p + (z * z) / (2 * n)) / d;
  const h = (z * Math.sqrt((p * (1 - p)) / n + (z * z) / (4 * n * n))) / d;
  return { k, n, rate: p, lo: Math.max(0, c - h), hi: Math.min(1, c + h) };
}
export const sha256 = async (buf) => (await import('node:crypto')).createHash('sha256').update(buf).digest('hex');
export function fmtRate(w) {
  if (!w.n) return 'n/a';
  const pct = (x) => (x * 100).toFixed(1) + '%';
  return `${w.k}/${w.n} (${pct(w.rate)}, 95% CI ${pct(w.lo)} to ${pct(w.hi)})`;
}
