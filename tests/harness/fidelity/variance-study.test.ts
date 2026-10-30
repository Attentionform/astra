/**
 * Renderer variance study (advisory only: nothing here blocks a merge).
 *
 * The study converts identical input bytes repeatedly (at least five runs)
 * and records run-to-run self-variance of the structural projection. The
 * numbers live in the tracked `variance-report.json` next to this file; the
 * live cases below re-measure on every run and log fresh numbers, but they
 * never fail on magnitude. What they do enforce:
 * - the projection pipeline is deterministic on identical bytes (re-parsing
 *   one PDF twice scores exactly zero, so any nonzero self-variance is
 *   attributable to the renderer, not the harness);
 * - the recorded report keeps a valid schema and its verdict obeys the
 *   promotion rule (a `promote` verdict requires at least three consecutive
 *   campaigns under budget with zero open converter crashers; a hand-edited
 *   promotion claim fails here).
 *
 * Promotion rule (recorded, not assumed): render similarity may gate merges
 * only after the study shows self-variance under budget for three
 * consecutive full campaigns AND every convertible fixture converts (open
 * crashers block promotion). Campaign history lives in the report file.
 */
import { describe, it, expect, afterAll } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  VARIANCE_BUDGET,
  createIsolatedProfile,
  isOracleAvailable,
  libreOfficeVersion,
  measureSelfVariance,
  parsePgm,
  pixelDifference,
  renderPdfToRasters,
  convertDocxToPdf,
} from './render-similarity';

const HERE = __dirname;
const FIX = path.join(HERE, 'fixtures');
const REPORT = path.join(HERE, 'variance-report.json');

const ORACLE = isOracleAvailable();
/** One renderer profile per test worker; see the advisory suite. */
const PROFILE = ORACLE ? createIsolatedProfile() : null;

afterAll(() => {
  if (PROFILE !== null) fs.rmSync(PROFILE, { recursive: true, force: true });
});

/**
 * One fixture per shape class (running prose, table grid, numbering part):
 * variance is a renderer property, so sampling shape classes covers the
 * space without paying six conversions per corpus fixture.
 */
const STUDY_INPUTS = ['gallery-coffee-onepager', 'merged-cells-basic', 'numbering-restart'];
const STUDY_RUNS = 6;

interface CampaignEntry {
  id: number;
  runsPerInput: number;
  inputs: Array<{
    fixture: string;
    maxSelfVariance: number;
    pagesStable: boolean;
    textStable: boolean;
  }>;
  maxSelfVariance: number;
}

interface VarianceReport {
  generatedAt: string;
  renderer: string;
  method: string;
  threshold: number;
  requiredConsecutiveCampaigns: number;
  campaigns: CampaignEntry[];
  openCrashers: string[];
  verdict: 'promote' | 'do-not-promote';
  reason: string;
}

function readReport(): VarianceReport {
  return JSON.parse(fs.readFileSync(REPORT, 'utf8')) as VarianceReport;
}

describe('variance study record', () => {
  it('variance report exists with a rule-consistent verdict', () => {
    expect(fs.existsSync(REPORT), 'variance-report.json must be recorded').toBe(true);
    const report = readReport();
    expect(typeof report.generatedAt).toBe('string');
    expect(typeof report.renderer).toBe('string');
    expect(report.threshold).toBe(VARIANCE_BUDGET);
    expect(report.requiredConsecutiveCampaigns).toBeGreaterThanOrEqual(3);
    expect(report.campaigns.length).toBeGreaterThanOrEqual(1);
    for (const campaign of report.campaigns) {
      expect(campaign.runsPerInput).toBeGreaterThanOrEqual(5);
      expect(campaign.inputs.length).toBeGreaterThanOrEqual(1);
      for (const input of campaign.inputs) {
        expect(Number.isFinite(input.maxSelfVariance)).toBe(true);
      }
      expect(Number.isFinite(campaign.maxSelfVariance)).toBe(true);
    }
    expect(['promote', 'do-not-promote']).toContain(report.verdict);
    expect(Array.isArray(report.openCrashers)).toBe(true);
    for (const crasher of report.openCrashers) expect(typeof crasher).toBe('string');
    if (report.verdict === 'promote') {
      // The promotion rule, enforced on the record: three consecutive
      // campaigns under budget AND zero open converter crashers, no
      // hand-waved promotion.
      const tail = report.campaigns.slice(-report.requiredConsecutiveCampaigns);
      expect(tail.length).toBe(report.requiredConsecutiveCampaigns);
      for (const campaign of tail) {
        expect(campaign.maxSelfVariance).toBeLessThan(report.threshold);
      }
      expect(report.openCrashers).toEqual([]);
      expect(report.reason).toMatch(/\S/);
    } else {
      expect(report.reason).toMatch(/campaign|convert|crash|variance/i);
    }
  });
});

describe.skipIf(!ORACLE)('renderer variance measurement', () => {
  it('projection of identical bytes is exactly deterministic', () => {
    const workRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-variance-det-'));
    try {
      const src = path.join(FIX, `${STUDY_INPUTS[0]}.docx`);
      const dir = fs.mkdtempSync(path.join(workRoot, 'one-'));
      const input = path.join(dir, 'input.docx');
      fs.copyFileSync(src, input);
      const pdf = convertDocxToPdf(input, dir, { profileDir: PROFILE as string });
      const first = renderPdfToRasters(pdf, dir);
      const second = renderPdfToRasters(pdf, dir);
      expect(first.length).toBe(second.length);
      for (let p = 0; p < first.length; p++) {
        expect(pixelDifference(first[p], second[p]).mean).toBe(0);
      }
      const reparsed = parsePgm(
        Buffer.concat([Buffer.from('P5\n1 1\n255\n', 'ascii'), Buffer.from(first[0].pixels.subarray(0, 1))]),
      );
      expect(reparsed.width).toBe(1);
    } finally {
      fs.rmSync(workRoot, { recursive: true, force: true });
    }
  }, 180_000);

  it('records run-to-run self-variance for identical inputs', () => {
    const workRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-variance-live-'));
    try {
      const summary: Array<{ fixture: string; maxSelfVariance: number; pagesStable: boolean; textStable: boolean }> = [];
      for (const stem of STUDY_INPUTS) {
        const result = measureSelfVariance(path.join(FIX, `${stem}.docx`), STUDY_RUNS, workRoot, PROFILE as string);
        expect(result.pages.length).toBe(STUDY_RUNS);
        expect(Number.isFinite(result.maxPixelMean)).toBe(true);
        summary.push({
          fixture: stem,
          maxSelfVariance: result.maxPixelMean,
          pagesStable: result.pagesStable,
          textStable: result.textStable,
        });
      }
      const worst = Math.max(...summary.map((s) => s.maxSelfVariance));
      console.log(
        `variance live (${libreOfficeVersion()}): ${JSON.stringify(summary)} worst=${worst} budget=${VARIANCE_BUDGET}`,
      );
    } finally {
      fs.rmSync(workRoot, { recursive: true, force: true });
    }
  }, 300_000);
});
