#!/usr/bin/env node
/**
 * Fidelity corpus manifest gate (no dependencies, plain node).
 *
 * Validates every tests/harness/fidelity/fixtures/*.manifest.json against
 * tests/harness/fidelity/manifest.schema.json. Fails on missing fields,
 * misnamed/unknown fields (including known banned aliases), word-count
 * violations, and malformed disposition blocks. Exit 0 when all pass.
 *
 * Usage: node scripts/check-fidelity-manifests.mjs [fixturesDir]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const schemaPath = path.join(root, 'tests/harness/fidelity/manifest.schema.json');
const fixturesDir = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(root, 'tests/harness/fidelity/fixtures');

const words = (s) => s.trim().split(/\s+/).filter(Boolean).length;

function checkManifest(manifest, stem, schema) {
  const errors = [];
  if (manifest === null || typeof manifest !== 'object' || Array.isArray(manifest)) {
    return ['manifest root must be a JSON object'];
  }
  for (const field of schema.requiredFields) {
    if (!(field in manifest)) errors.push(`missing required field '${field}'`);
  }
  const allowed = new Set([...schema.requiredFields, ...schema.optionalFields]);
  for (const key of Object.keys(manifest)) {
    if (!allowed.has(key)) {
      if (schema.bannedAliases.includes(key)) {
        errors.push(`banned alias '${key}' is not a schema field (misnamed; use the canonical name)`);
      } else {
        errors.push(`unknown field '${key}' (misnamed or out of schema; no additional properties allowed)`);
      }
    }
  }
  if ('fixture' in manifest) {
    if (manifest.fixture !== stem) errors.push(`fixture '${manifest.fixture}' does not match file stem '${stem}'`);
    if (typeof manifest.fixture !== 'string' || !new RegExp(schema.namePattern).test(manifest.fixture)) {
      errors.push('fixture must match ' + schema.namePattern);
    }
  }
  if ('family' in manifest && (typeof manifest.family !== 'string' || !new RegExp(schema.namePattern).test(manifest.family))) {
    errors.push('family must match ' + schema.namePattern);
  }
  for (const [field, [lo, hi]] of Object.entries(schema.wordRanges)) {
    if (field.startsWith('disposition.')) continue;
    if (field in manifest) {
      if (typeof manifest[field] !== 'string') errors.push(`'${field}' must be a string`);
      else {
        const n = words(manifest[field]);
        if (n < lo || n > hi) errors.push(`'${field}' has ${n} words, required ${lo}-${hi}`);
      }
    }
  }
  if ('disposition' in manifest) {
    const d = manifest.disposition;
    if (d === null || typeof d !== 'object' || Array.isArray(d)) {
      errors.push('disposition must be an object {verdict, reason}');
    } else {
      const dKeys = new Set(Object.keys(d));
      if (!dKeys.has('verdict') || !dKeys.has('reason')) errors.push('disposition requires verdict and reason');
      for (const k of dKeys) {
        if (k !== 'verdict' && k !== 'reason') errors.push(`disposition has unknown field '${k}'`);
      }
      if ('verdict' in d && !schema.dispositionVerdicts.includes(d.verdict)) {
        errors.push(`disposition.verdict must be one of ${schema.dispositionVerdicts.join('|')}`);
      }
      if ('reason' in d) {
        const [lo, hi] = schema.wordRanges['disposition.reason'];
        if (typeof d.reason !== 'string') errors.push('disposition.reason must be a string');
        else {
          const n = words(d.reason);
          if (n < lo || n > hi) errors.push(`disposition.reason has ${n} words, required ${lo}-${hi}`);
        }
      }
    }
  }
  if ('sourceDocument' in manifest && typeof manifest.sourceDocument !== 'string') {
    errors.push('sourceDocument must be a string');
  }
  return errors;
}

function main() {
  const schema = JSON.parse(fs.readFileSync(schemaPath, 'utf8'));
  const files = fs.readdirSync(fixturesDir).filter((f) => f.endsWith('.manifest.json')).sort();
  if (files.length === 0) {
    console.error(`no manifests found in ${fixturesDir}`);
    process.exit(1);
  }
  let failures = 0;
  for (const file of files) {
    const stem = file.replace(/\.manifest\.json$/, '');
    const full = path.join(fixturesDir, file);
    let manifest;
    try {
      manifest = JSON.parse(fs.readFileSync(full, 'utf8'));
    } catch (err) {
      console.error(`FAIL ${file}: invalid JSON (${err.message})`);
      failures++;
      continue;
    }
    const errors = checkManifest(manifest, stem, schema);
    const docxPath = path.join(fixturesDir, `${stem}.docx`);
    if (!fs.existsSync(docxPath)) errors.push(`missing sibling document '${stem}.docx'`);
    if (errors.length > 0) {
      failures++;
      for (const e of errors) console.error(`FAIL ${file}: ${e}`);
    } else {
      console.log(`ok ${file}`);
    }
  }
  if (failures > 0) {
    console.error(`${failures} manifest(s) failed`);
    process.exit(1);
  }
  console.log(`${files.length} manifest(s) pass`);
}

main();
