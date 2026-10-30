import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';

describe('Byte-Hash Fidelity Automated Test (Zero-Touch Invariance)', () => {
  it('preserves checksums of all untouched nodes during surgical mutation on T1_coffee_onepager.docx', async () => {
    const docPath = path.resolve(__dirname, '../../test-gallery/T1_coffee_onepager.docx');
    const buf = fs.readFileSync(docPath);
    const parser = new DocxParser(buf);
    const initialAst = await parser.getAST();

    expect(initialAst.sections.length).toBeGreaterThanOrEqual(1);
    const section = initialAst.sections[0];
    expect(section.nodes.length).toBeGreaterThan(2);

    // 1. Capture initial node checksums and raw XML
    const initialChecksums = new Map<string, string>();
    const initialXmls = new Map<string, string>();
    for (const sec of initialAst.sections) {
      for (const node of sec.nodes) {
        expect(node.checksum).toBeTruthy();
        initialChecksums.set(node.id, node.checksum);
        if (node.xml) {
          initialXmls.set(node.id, node.xml.toString());
        }
      }
    }

    // 2. Target a single paragraph for surgical mutation
    const targetNode = section.nodes[1];
    const initialTargetChecksum = targetNode.checksum;
    const initialTargetText = targetNode.metadata.text;

    const engine = new OperationEngine(initialAst, parser);
    const batchResult = await engine.executeBatch({
      operations: [
        {
          action: 'replace',
          target: targetNode.id,
          text: 'Artisan single-origin Ethiopian roast updated surgically with zero touch side effects.',
        },
      ],
      sequential: true,
    });
    expect(batchResult.success).toBe(true);

    // 3. Re-zip and re-parse completely from disk bytes
    const mutatedBuf = parser.rezip();
    const reloadedParser = new DocxParser(mutatedBuf);
    const reloadedAst = await reloadedParser.getAST();

    // 4. Assert fidelity invariants
    let foundTarget = false;
    let untouchedCount = 0;

    for (const sec of reloadedAst.sections) {
      for (const node of sec.nodes) {
        if (node.id === targetNode.id) {
          foundTarget = true;
          // Target checksum must change
          expect(node.checksum).not.toBe(initialTargetChecksum);
          expect(node.metadata.text).not.toBe(initialTargetText);
        } else {
          // Untouched node checksum must be 100% byte-for-byte identical
          const expectedChecksum = initialChecksums.get(node.id);
          expect(expectedChecksum).toBeDefined();
          expect(node.checksum).toBe(expectedChecksum);
          untouchedCount++;
        }
      }
    }

    expect(foundTarget).toBe(true);
    expect(untouchedCount).toBe(initialChecksums.size - 1);
  });

  it('preserves untouched sections and untouched nodes across multi-section document T2_edit_target.docx', async () => {
    const docPath = path.resolve(__dirname, '../../test-gallery/T2_edit_target.docx');
    const buf = fs.readFileSync(docPath);
    const parser = new DocxParser(buf);
    const initialAst = await parser.getAST();

    const initialChecksums = new Map<string, string>();
    const sectionChecksums = new Map<string, string>();
    for (const sec of initialAst.sections) {
      sectionChecksums.set(sec.id, sec.checksum);
      for (const node of sec.nodes) {
        initialChecksums.set(node.id, node.checksum);
      }
    }

    // Pick a target node in the last section
    const targetSection = initialAst.sections[initialAst.sections.length - 1];
    expect(targetSection.nodes.length).toBeGreaterThanOrEqual(1);
    const targetNode = targetSection.nodes[0];
    const initialTargetChecksum = targetNode.checksum;

    const engine = new OperationEngine(initialAst, parser);
    const res = await engine.executeBatch({
      operations: [
        {
          action: 'replace',
          target: targetNode.id,
          text: 'Verified Zero-Touch Invariance Replacement.',
        },
      ],
      sequential: true,
    });
    expect(res.success).toBe(true);

    const mutatedBuf = parser.rezip();
    const reloadedParser = new DocxParser(mutatedBuf);
    const reloadedAst = await reloadedParser.getAST();

    // Verify all earlier untouched sections remain 100% byte-identical
    for (let s = 0; s < reloadedAst.sections.length - 1; s++) {
      const untouchedSec = reloadedAst.sections[s];
      expect(untouchedSec.checksum).toBe(sectionChecksums.get(untouchedSec.id));
      for (const node of untouchedSec.nodes) {
        expect(node.checksum).toBe(initialChecksums.get(node.id));
      }
    }

    // In the modified section, all other nodes remain identical
    const modifiedSec = reloadedAst.sections[reloadedAst.sections.length - 1];
    for (const node of modifiedSec.nodes) {
      if (node.id === targetNode.id) {
        expect(node.checksum).not.toBe(initialTargetChecksum);
      } else {
        expect(node.checksum).toBe(initialChecksums.get(node.id));
      }
    }
  });
});
