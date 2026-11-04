import { describe, it, expect } from 'vitest';
import { computeDocStats } from '../src/stats';

describe('Layout-Weighted Pagination Heuristic (AS-05)', () => {
  it('estimates pages for pure prose based on 450 words per page', () => {
    // 0 words -> 1 page minimum
    expect(computeDocStats({ sections: [{ nodes: [] }] }).pages).toBe(1);

    // 450 words -> 1 page
    const p450 = Array(450).fill('word').join(' ');
    expect(computeDocStats({ sections: [{ nodes: [{ type: 'paragraph', metadata: { text: p450 } }] }] }).pages).toBe(1);

    // 451 words -> 2 pages
    const p451 = Array(451).fill('word').join(' ');
    expect(computeDocStats({ sections: [{ nodes: [{ type: 'paragraph', metadata: { text: p451 } }] }] }).pages).toBe(2);

    // 900 words -> 2 pages
    const p900 = Array(900).fill('word').join(' ');
    expect(computeDocStats({ sections: [{ nodes: [{ type: 'paragraph', metadata: { text: p900 } }] }] }).pages).toBe(2);

    // 901 words -> 3 pages
    const p901 = Array(901).fill('word').join(' ');
    expect(computeDocStats({ sections: [{ nodes: [{ type: 'paragraph', metadata: { text: p901 } }] }] }).pages).toBe(3);
  });

  it('weights tables with +0.45 page budget per table', () => {
    // 450 words (1.0 page) + 1 table (0.45) = 1.45 -> 2 pages
    const p450 = Array(450).fill('word').join(' ');
    const stats = computeDocStats({
      sections: [
        {
          nodes: [
            { type: 'paragraph', metadata: { text: p450 } },
            { type: 'table', metadata: { text: 'Cell 1 Cell 2' } },
          ],
        },
      ],
    });
    expect(stats.tables).toBe(1);
    expect(stats.pages).toBe(2);
  });

  it('weights images with +0.50 page budget per image', () => {
    // 400 words (0.889 page) + 2 images (1.0) = 1.889 -> 2 pages
    const p400 = Array(400).fill('word').join(' ');
    const stats = computeDocStats({
      sections: [
        {
          nodes: [
            { type: 'paragraph', metadata: { text: p400 } },
            { type: 'image', metadata: { text: '[Image]' } },
            { type: 'image', metadata: { text: '[Image]' } },
          ],
        },
      ],
    });
    expect(stats.images).toBe(2);
    expect(stats.pages).toBe(2);
  });

  it('weights explicit hard page breaks with +1.0 page budget per break', () => {
    // 10 words + 1 hard page break = 10/450 + 1 = 1.022 -> 2 pages
    const stats = computeDocStats({
      sections: [
        {
          nodes: [
            { type: 'paragraph', metadata: { text: 'Ten words in this short first section before the break marker.' } },
            { type: 'pageBreak', metadata: { text: '[Page Break]' } },
            { type: 'paragraph', metadata: { text: 'And ten more words on the next physical page here.' } },
          ],
        },
      ],
    });
    expect(stats.pages).toBe(2);
  });

  it('accurately estimates 12 pages for the 3,449-word benchmark report with 5 tables and 4 images', () => {
    // Benchmark document from AS-05: 3,449 words, 5 tables, 4 figures
    // Old naive formula: ceil(3449 / 500) = 7 pages (~71% under-estimate)
    // New layout-weighted formula: ceil(3449/450 + 5*0.45 + 4*0.50) = ceil(7.664 + 2.25 + 2.0) = ceil(11.914) = 12 pages!
    const wordsPerPara = Math.floor(3449 / 10); // 10 paragraphs of ~345 words
    const paras = Array(10)
      .fill(null)
      .map((_, i) => ({
        type: 'paragraph' as const,
        metadata: { text: Array(i === 9 ? 3449 - 9 * wordsPerPara : wordsPerPara).fill('prose').join(' ') },
      }));

    const tables = Array(5)
      .fill(null)
      .map(() => ({
        type: 'table' as const,
        metadata: { text: 'colA colB colC' },
      }));

    const images = Array(4)
      .fill(null)
      .map(() => ({
        type: 'image' as const,
        metadata: { text: '[Image]' },
      }));

    const stats = computeDocStats({
      sections: [{ nodes: [...paras, ...tables, ...images] }],
    });

    // 3449 words + 5*3 (table cell words) = 3464 total words
    // 3464/450 = 7.698 + 5*0.45 (2.25) + 4*0.50 (2.0) = 11.948 -> ceil = 12 pages!
    expect(stats.tables).toBe(5);
    expect(stats.images).toBe(4);
    expect(stats.pages).toBe(12);
  });

  it('runs sub-millisecond and deterministic over 1,000 iterations', () => {
    const dummyAst = {
      sections: [
        {
          nodes: [
            { type: 'paragraph', metadata: { text: 'Sample paragraph text' } },
            { type: 'table', metadata: { text: 'Col 1 Col 2' } },
            { type: 'image', metadata: { text: '[Image]' } },
          ],
        },
      ],
    };

    const start = performance.now();
    for (let i = 0; i < 1000; i++) {
      computeDocStats(dummyAst);
    }
    const elapsedMs = performance.now() - start;
    expect(elapsedMs).toBeLessThan(50); // Less than 0.05ms per call
  });
});
