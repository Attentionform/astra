import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  assertEngineQuotas,
  DocxParser,
  ENGINE_MAX_BYTES,
  ENGINE_MAX_OPS,
  ENGINE_MAX_STEPS,
  OperationEngine,
  validateNode,
  validateOperationsStrict,
  validateOperationStrict,
} from '@attentionform/astra-engine';

/**
 * Sprint-3 chaos soak (50x randomized generation + mutation).
 *
 * - Mixed ops every iteration: replace/insert batch, table, style, redline
 *   (trackedChange), plus rotating list/footnote/field coverage.
 * - Mixed execution modes: sequential true/false alternate (document order vs
 *   legacy sort-order path), single + multi-op batches.
 * - Bounded: ops <= 500, JSON bytes <= 200MB; maxSteps is an optional positive
 *   integer (unbounded when omitted; ENGINE_MAX_STEPS=25 is a documented hint
 *   only, not a ceiling — asserted per batch via assertEngineQuotas; maxima
 *   reported at the end).
 * - Deterministic: seeded PRNG (mulberry32). Override with CHAOS_SEED=<int>
 *   to replay an exact sequence; the seed is logged on every run.
 * - Asserts: 0 OOM, 0 unhandled rejections/exceptions, 100% validator
 *   ok:true (strict schema + quotas + 6-layer validateNode spot-checks +
 *   package integrity), every batch applied without rollback.
 */

const SOAK_ITERATIONS = 50;

function resolveSeed(): number {
  const raw = (process.env.CHAOS_SEED || '').trim();
  if (/^-?\d+$/.test(raw)) return Number.parseInt(raw, 10) >>> 0;
  return 0xc0ffee;
}

function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function createSoakDocBuffer(paragraphs = 8): Buffer {
  const bodyParas: string[] = [];
  bodyParas.push(
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Soak Heading Alpha</w:t></w:r></w:p>`
  );
  for (let i = 0; i < paragraphs; i++) {
    bodyParas.push(
      `<w:p><w:r><w:t>Soak body paragraph ${i} with stable content for mutation.</w:t></w:r></w:p>`
    );
  }
  bodyParas.push(
    `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Soak Heading Beta</w:t></w:r></w:p>`
  );
  bodyParas.push(
    `<w:p><w:r><w:t>Soak tail paragraph closing the section.</w:t></w:r></w:p>`
  );
  const tableCells = (row: number): string => {
    let tcs = '';
    for (let c = 0; c < 3; c++) {
      tcs += `<w:tc><w:p><w:r><w:t>R${row}C${c}</w:t></w:r></w:p></w:tc>`;
    }
    return `<w:tr>${tcs}</w:tr>`;
  };
  const docXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body>${bodyParas.join('')}` +
    `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3360"/></w:tblGrid>` +
    `${tableCells(0)}${tableCells(1)}${tableCells(2)}</w:tbl>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>` +
    `</w:body></w:document>`;
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `</Types>`;
  const zip = new AdmZip();
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from(contentTypes, 'utf8'));
  return zip.toBuffer();
}

describe('Sprint-3 chaos soak: 50x randomized generation+mutation', () => {
  it(
    'completes 50 iterations with 0 OOM, 0 unhandled, 100% validator ok:true (bounded 500/200MB, maxSteps unbounded)',
    async () => {
      const seed = resolveSeed();
      // eslint-disable-next-line no-console
      console.log(`[chaos-soak] seed=${seed} iterations=${SOAK_ITERATIONS}`);
      const rng = mulberry32(seed);
      const pick = <T>(arr: readonly T[]): T => arr[Math.floor(rng() * arr.length)];
      const pickIndex = (len: number): number => Math.floor(rng() * len);

      const unhandled: unknown[] = [];
      const onUnhandled = (e: unknown): void => {
        unhandled.push(e);
      };
      process.on('unhandledRejection', onUnhandled);
      process.on('uncaughtException', onUnhandled);
      try {
        let oomCount = 0;
        let expectedRejects = 0;
        let validatorOk = 0;
        let validatorTotal = 0;
        let successBatches = 0;
        let batchCount = 0;
        let sequentialCount = 0;
        let sortedCount = 0;
        let tableOps = 0;
        let styleOps = 0;
        let redlineOps = 0;
        let maxOpsSeen = 0;
        let maxBytesSeen = 0;
        let maxStepsSeen = 0;

        for (let iter = 0; iter < SOAK_ITERATIONS; iter++) {
          const parser = new DocxParser(createSoakDocBuffer(6 + (iter % 4)));
          const ast = await parser.getAST();
          const engine = new OperationEngine(ast, parser);
          const paraIds = ast.sections.flatMap((s) =>
            s.nodes.filter((n) => n.type === 'paragraph').map((n) => n.id)
          );
          const tableIds = ast.sections.flatMap((s) =>
            s.nodes.filter((n) => n.type === 'table').map((n) => n.id)
          );
          expect(paraIds.length).toBeGreaterThan(3);
          expect(tableIds.length).toBeGreaterThan(0);

          const sequential = iter % 2 === 0;
          // ENGINE_MAX_STEPS is a hint only (not a ceiling): most iterations draw
          // 1..hint, every 10th probes freedom (200/1000 must also pass).
          const maxSteps = iter % 10 === 9 ? (iter % 20 === 9 ? 200 : 1000) : 1 + Math.floor(rng() * ENGINE_MAX_STEPS);
          const ops: Record<string, unknown>[] = [];
          const usedTargets = new Set<string>();
          const takePara = (): string => {
            for (let tries = 0; tries < 8; tries++) {
              const id = paraIds[pickIndex(paraIds.length)];
              if (!usedTargets.has(id)) {
                usedTargets.add(id);
                return id;
              }
            }
            return paraIds[pickIndex(paraIds.length)];
          };
          const tableId = tableIds[pickIndex(tableIds.length)];

          // 1. replace (generation baseline — also the validateNode spot-check target)
          const replaceTarget = takePara();
          let replaceBefore = '';
          try {
            replaceBefore = parser.readNode(replaceTarget);
          } catch {
            replaceBefore = '';
          }
          ops.push({
            action: 'replace',
            target: replaceTarget,
            text: `Soak iter ${iter} replacement text with seed ${seed}.`,
          });

          // 2. style (required coverage)
          ops.push({
            action: 'style',
            target: takePara(),
            properties: {
              bold: rng() < 0.5,
              italic: rng() < 0.5,
              color: pick(['1D4ED8', '0F766E', 'B45309'] as const),
            },
          });

          // 3. table (required coverage — index-safe sub-ops only)
          ops.push(
            pick([
              {
                action: 'table',
                target: tableId,
                operation: 'updateCell',
                rowIndex: 1,
                colIndex: 1,
                text: `S${iter}`,
              },
              { action: 'table', target: tableId, operation: 'setShading', fill: 'F7F9FC' },
              {
                action: 'table',
                target: tableId,
                operation: 'setBorder',
                border: 'all',
                color: 'D1D5DB',
                sz: 4,
              },
            ])
          );

          // 4. redline (required coverage — trackedChange insert/delete)
          ops.push(
            rng() < 0.6
              ? {
                  action: 'trackedChange',
                  operation: 'insert',
                  target: takePara(),
                  text: `redline iter ${iter}`,
                  author: 'SoakBot',
                }
              : { action: 'trackedChange', operation: 'delete', target: takePara(), author: 'SoakBot' }
          );

          // 5-6. rotating extras: insert / list / footnote / field
          const extras = 1 + Math.floor(rng() * 2);
          for (let e = 0; e < extras; e++) {
            const kind = pick(['insert', 'list', 'footnote', 'field'] as const);
            if (kind === 'insert') {
              ops.push(
                rng() < 0.75
                  ? {
                      action: 'insert',
                      type: 'paragraph',
                      anchor: paraIds[pickIndex(paraIds.length)],
                      position: rng() < 0.5 ? 'before' : 'after',
                      text: `Soak insert iter ${iter}.${e}`,
                    }
                  : {
                      action: 'insert',
                      type: 'table',
                      anchor: paraIds[pickIndex(paraIds.length)],
                      position: 'after',
                      headers: ['H1', 'H2'],
                      rows: [
                        [`a${iter}`, `b${iter}`],
                        [`c${iter}`, `d${iter}`],
                      ],
                    }
              );
            } else if (kind === 'list') {
              ops.push({
                action: 'list',
                operation: pick(['create', 'continue', 'changeLevel'] as const),
                target: takePara(),
                listType: rng() < 0.5 ? 'bullet' : 'numbered',
              });
            } else if (kind === 'footnote') {
              ops.push({
                action: 'footnote',
                operation: 'insert',
                anchor: takePara(),
                text: `Soak footnote iter ${iter}.${e}`,
              });
            } else {
              const fop = pick(['insertPageNumber', 'insertWordCount', 'updateFields'] as const);
              ops.push(
                fop === 'updateFields'
                  ? { action: 'field', operation: 'updateFields' }
                  : {
                      action: 'field',
                      operation: fop,
                      anchor: paraIds[pickIndex(paraIds.length)],
                      position: 'after',
                    }
              );
            }
          }

          // Bounded quotas (fail-closed before any mutation).
          const bytes = Buffer.byteLength(JSON.stringify(ops), 'utf8');
          maxOpsSeen = Math.max(maxOpsSeen, ops.length);
          maxBytesSeen = Math.max(maxBytesSeen, bytes);
          maxStepsSeen = Math.max(maxStepsSeen, maxSteps);
          expect(ops.length).toBeLessThanOrEqual(ENGINE_MAX_OPS);
          expect(bytes).toBeLessThanOrEqual(ENGINE_MAX_BYTES);
          expect(Number.isInteger(maxSteps) && maxSteps >= 1).toBe(true);

          // Strict schema validators — every op and the batch must be ok:true.
          for (const op of ops) {
            validatorTotal += 1;
            expect(() => validateOperationStrict(op)).not.toThrow();
            validatorOk += 1;
          }
          validatorTotal += 1;
          expect(() => validateOperationsStrict(ops)).not.toThrow();
          validatorOk += 1;
          validatorTotal += 1;
          expect(() => assertEngineQuotas(ops as never, maxSteps)).not.toThrow();
          validatorOk += 1;

          let res;
          let iterSucceeded = false;
          try {
            res = await engine.executeBatch({
              operations: ops as never,
              sequential,
              maxSteps,
            });
          } catch (err: unknown) {
            const msg = err instanceof Error ? err.message + (err.stack || '') : String(err);
            if (/out of memory|heap|allocation failed/i.test(msg)) oomCount++;
            throw err;
          }
          batchCount++;
          if (sequential) sequentialCount++;
          else sortedCount++;
          tableOps += ops.filter((o) => o['action'] === 'table').length;
          styleOps += ops.filter((o) => o['action'] === 'style').length;
          redlineOps += ops.filter((o) => o['action'] === 'trackedChange').length;

          if (!res.success) {
            // Designed fail-closed: a random `continue` with no list behind
            // it rejects loudly instead of forking a silent sequence. Only
            // that rejection is accepted here — anything else still fails.
            const errs = (res.errors || [])
              .map((e) => String((e as { error?: unknown }).error ?? e))
              .join('\n');
            const hasContinue = ops.some((o) => o['action'] === 'list' && o['operation'] === 'continue');
            expect(hasContinue, `batch failed without a continue op: ${errs.slice(0, 300)}`).toBe(true);
            expect(errs).toMatch(/no previous list item/);
            expectedRejects++;
          } else {
            expect(res.rollbackTriggered).toBe(false);
            iterSucceeded = true;
          }

          // 6-layer validator spot-check on the replace target (style/run
          // drift allowed — structure/layer preservation is under test).
          // Loud assert: the baseline must exist — never silently skip the
          // spot-check when readNode fails or returns empty.
          expect(replaceBefore).toBeTruthy();
          if (replaceBefore) {
            try {
              const after = parser.readNode(replaceTarget);
              validatorTotal += 1;
              expect(
                validateNode(replaceBefore, after, {
                  allowStyleChanges: true,
                  allowRunCountChanges: true,
                })
              ).toBe(true);
              validatorOk += 1;
            } catch (err: unknown) {
              const msg = err instanceof Error ? err.message : String(err);
              if (/out of memory|heap|allocation failed/i.test(msg)) oomCount++;
              throw err;
            }
          }

          // Package integrity + export still valid after the mutation storm.
          validatorTotal += 1;
          expect(parser.validateZipIntegrity().isValid).toBe(true);
          validatorOk += 1;
          const exported = parser.rezip();
          expect(exported.length).toBeGreaterThan(1000);
          if (iterSucceeded) successBatches++;
        }

        // Soak-level assertions.
        expect(successBatches + expectedRejects).toBe(SOAK_ITERATIONS);
        expect(batchCount).toBe(SOAK_ITERATIONS);
        expect(sequentialCount).toBeGreaterThan(0);
        expect(sortedCount).toBeGreaterThan(0);
        expect(tableOps).toBeGreaterThan(0);
        expect(styleOps).toBeGreaterThan(0);
        expect(redlineOps).toBeGreaterThan(0);
        expect(oomCount).toBe(0);
        expect(unhandled.length).toBe(0);
        expect(validatorTotal).toBeGreaterThan(0);
        expect(validatorOk).toBe(validatorTotal);
        expect(maxOpsSeen).toBeLessThanOrEqual(ENGINE_MAX_OPS);
        expect(maxBytesSeen).toBeLessThanOrEqual(ENGINE_MAX_BYTES);
        expect(maxStepsSeen).toBeGreaterThanOrEqual(1);
        // eslint-disable-next-line no-console
        console.log(
          `[chaos-soak] done seed=${seed} batches=${batchCount} ` +
            `table=${tableOps} style=${styleOps} redline=${redlineOps} ` +
            `validator=${validatorOk}/${validatorTotal} maxOps=${maxOpsSeen} maxBytes=${maxBytesSeen} maxSteps=${maxStepsSeen}`
        );
      } finally {
        process.off('unhandledRejection', onUnhandled);
        process.off('uncaughtException', onUnhandled);
      }
    },
    120000
  );
});
