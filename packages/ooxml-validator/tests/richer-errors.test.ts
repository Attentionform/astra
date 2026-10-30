import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { validateFile, validateBuffer } from '../src/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES_DIR = path.join(__dirname, 'fixtures');
const INVALID_SCHEMA_DOCX = path.join(FIXTURES_DIR, 'invalid-schema.docx');
const MANY_ERRORS_DOCX = path.join(FIXTURES_DIR, 'many-errors.docx');
const VALID_DOCX = path.join(FIXTURES_DIR, 'valid.docx');

describe('Richer Error Attribution (Milestone M4)', () => {
  it('schema error returns node.localName, part, class=schema, and relatedNode', async () => {
    const bytes = fs.readFileSync(INVALID_SCHEMA_DOCX);
    const res = await validateBuffer(bytes, { officeVersion: 'Office2019' });

    expect(res.ok).toBe(false);
    expect(res.errors.length).toBeGreaterThan(0);

    const err = res.errors[0];
    expect(err.errorType).toBe('Schema');
    expect(err.class).toBe('schema');
    expect(err.part).toBe('/word/document.xml');
    expect(err.path).toBe('/word/document.xml');
    expect(err.node).toBeDefined();
    expect(err.node?.localName).toBe('body');
    expect(err.node?.namespaceUri).toBe(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    );
    expect(err.relatedNode).toBeDefined();
    expect(err.relatedNode?.localName).toBe('bogusElementWithInvalidSchema');
  });

  it('validateFile also returns richer node, part, and class attributes', async () => {
    const res = await validateFile(INVALID_SCHEMA_DOCX, { officeVersion: 'Office2019' });
    expect(res.ok).toBe(false);
    expect(res.errors.length).toBeGreaterThan(0);

    const err = res.errors[0];
    expect(err.class).toBe('schema');
    expect(err.part).toBe('/word/document.xml');
    expect(err.node?.localName).toBe('body');
    expect(err.relatedNode?.localName).toBe('bogusElementWithInvalidSchema');
  });

  it('semantic error is classed as semantic', async () => {
    const baseBytes = fs.readFileSync(VALID_DOCX);
    const zip = new AdmZip(baseBytes);
    let xml = zip.readAsText('word/document.xml');
    xml = xml.replace(
      '</w:body>',
      '<w:p><w:bookmarkStart w:id="0" w:name="b1"/><w:bookmarkStart w:id="0" w:name="b2"/></w:p></w:body>',
    );
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const modifiedBytes = zip.toBuffer();

    const res = await validateBuffer(modifiedBytes, { officeVersion: 'Office2019' });
    expect(res.ok).toBe(false);

    const semErr = res.errors.find(
      (e) => e.class === 'semantic' || e.errorType === 'Semantic' || e.id?.startsWith('Sem_'),
    );
    expect(semErr).toBeDefined();
    expect(semErr?.class).toBe('semantic');
    expect(semErr?.errorType).toBe('Semantic');
    expect(semErr?.node?.localName).toBe('bookmarkStart');
  });

  it('richer fields are present on every error of a multi-error document', async () => {
    const bytes = fs.readFileSync(MANY_ERRORS_DOCX);
    const res = await validateBuffer(bytes, { officeVersion: 'Office2007', maxErrors: 15 });

    expect(res.ok).toBe(false);
    expect(res.errors.length).toBe(15);

    for (const err of res.errors) {
      expect(err.class).toBeTruthy();
      expect(err.part).toBe('/word/document.xml');
      expect(err.node).toBeDefined();
      expect(typeof err.node?.localName).toBe('string');
      expect(typeof err.node?.namespaceUri).toBe('string');
    }
  });
});
