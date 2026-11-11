// Seeded Markdown generator for C1. Stays inside the documented compose_document subset.
import { seededRng } from './exp.mjs';
const V = 'account agreement analysis annual approach asset audit balance baseline benefit board budget capacity capital client clause committee compliance contract control cost customer data deadline delivery demand design driver effect efficiency employee estimate evidence expense factor forecast framework funding goal governance growth impact income indicator initiative insight investment issue launch liability liquidity margin market measure milestone objective obligation operation outcome partner performance pipeline plan policy portfolio priority process procurement product profit project proposal quality quarter record reduction regulation requirement reserve resource result review revenue risk roadmap scenario schedule scope segment service signal stakeholder standard strategy summary supplier target team threshold timeline trend update utilization valuation variance vendor volume workflow'.split(' ');
export function genMarkdown(pages, seed) {
  const r = seededRng(seed);
  const w = () => V[Math.floor(r() * V.length)];
  const words = (n) => Array.from({ length: n }, w);
  const sent = (lo = 8, hi = 16) => { const x = words(lo + Math.floor(r() * (hi - lo))); x[0] = x[0][0].toUpperCase() + x.slice(0).shift().slice(1); return x.join(' ') + '.'; };
  const blocks = []; const md = []; let fn = 0;
  const target = Math.max(4, Math.round(pages * 6));
  for (let i = 0; i < target; i++) {
    const roll = r();
    if (i === 0) { const t = sent(3, 6).replace(/\.$/, ''); blocks.push({ kind: 'h', level: 1, text: t }); md.push(`# ${t}\n`); continue; }
    if (roll < 0.12) { const level = 2 + Math.floor(r() * 2); const t = sent(2, 5).replace(/\.$/, ''); blocks.push({ kind: 'h', level, text: t }); md.push(`${'#'.repeat(level)} ${t}\n`); }
    else if (roll < 0.52) {
      const a = words(3).join(' '), b = w(), c = w(), d = w();
      const plain = `${sent()} ${a} ${b} ${c} ${sent(5, 9)}`; // text without markers
      const text = `${sent()} **${a}** *${b}* \`${c}\` [${d}](https://example.org/${d}) ${sent(5, 9)}`;
      const expected = text.replace(/\*\*/g, '').replace(/\*/g, '').replace(/`/g, '').replace(/\[([^\]]+)\]\([^)]+\)/g, '$1');
      let t2 = expected, md2 = text;
      if (r() < 0.1) { fn++; md2 += `[^${fn}]`; blocks.push({ kind: 'p', text: t2 }); md.push(md2 + '\n'); const ft = sent(); md.push(`[^${fn}]: ${ft}\n`); blocks.push({ kind: 'fn', label: String(fn), text: ft }); continue; }
      blocks.push({ kind: 'p', text: t2 }); md.push(md2 + '\n');
    } else if (roll < 0.64) { const items = Array.from({ length: 3 + Math.floor(r() * 4) }, () => sent(4, 9)); blocks.push({ kind: 'ul', items }); md.push(items.map((x) => `- ${x}`).join('\n') + '\n'); }
    else if (roll < 0.74) { const items = Array.from({ length: 3 + Math.floor(r() * 3) }, () => sent(4, 9)); blocks.push({ kind: 'ol', items }); md.push(items.map((x, k) => `${k + 1}. ${x}`).join('\n') + '\n'); }
    else if (roll < 0.86) {
      const cols = 3 + Math.floor(r() * 3), nr = 3 + Math.floor(r() * 5);
      const rows = [Array.from({ length: cols }, () => w())]; for (let k = 1; k < nr; k++) rows.push(Array.from({ length: cols }, () => `${w()} ${Math.floor(r() * 9999)}`));
      blocks.push({ kind: 'table', rows });
      md.push(`| ${rows[0].join(' | ')} |\n| ${rows[0].map(() => '---').join(' | ')} |\n` + rows.slice(1).map((x) => `| ${x.join(' | ')} |`).join('\n') + '\n');
    } else if (roll < 0.92) { const t = sent(8, 14); blocks.push({ kind: 'quote', text: t }); md.push(`> ${t}\n`); }
    else if (roll < 0.97) { const lines = Array.from({ length: 4 }, () => `${w()}_${w()} = ${Math.floor(r() * 99)}`); blocks.push({ kind: 'code', lines }); md.push('```\n' + lines.join('\n') + '\n```\n'); }
    else { blocks.push({ kind: 'hr' }); md.push('---\n'); }
  }
  return { markdown: md.join('\n'), blocks };
}
