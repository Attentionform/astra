import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { latexToOmml, ommlToLatex } from '../../packages/engine/src/ooxml/omml';

const NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const doc = new DOMParser().parseFromString(`<w:p ${NS}/>`, 'text/xml') as unknown as Document;
const ser = (el: Element): string => new XMLSerializer().serializeToString(el as never);

/** latex -> omml1 -> latex2 -> omml2: the two OMML trees must serialize the same. */
function roundTrip(latex: string): { latex2: string | null; same: boolean } {
  const omml1 = latexToOmml(doc, latex);
  const latex2 = ommlToLatex(omml1);
  if (latex2 === null) return { latex2, same: false };
  return { latex2, same: ser(omml1) === ser(latexToOmml(doc, latex2)) };
}

const FORMULAS = [
  // from equation-omml-e4.test.ts and equation-latex-omml.test.ts
  'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
  '\\sqrt{a^2 + b^2} = c,\\ \\sqrt[3]{x}',
  'Q_{\\alpha}^{2} + x_{i}',
  '\\sum_{i=1}^{n} a_i',
  '\\int_{a}^{b} f(x) dx',
  '\\left( \\frac{a}{b} \\right) + |c|',
  '\\nabla \\cdot \\alpha \\le \\beta \\ge \\gamma \\neq \\pi \\approx \\infty',
  '\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1',
  '\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J} + \\mu_0 \\epsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t}',
  '\\frac{x + \\frac{1}{y}}{2}',
  'E = \\frac{m}{\\sqrt{c}}',
  '\\text{a b} + \\mbox{c d}',
  '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}',
  '\\begin{matrix} a & b & c \\end{matrix}',
  '\\begin{bmatrix} a & b & c \\\\ d & e & f \\end{bmatrix}',
  '\\begin{array}{cc} a & b \\\\ c & d \\end{array}',
  '\\begin{vmatrix} a \\\\ b \\\\ c \\end{vmatrix}',
  '\\alpha x^2y', '\\alpha x_ij', '\\alpha e^ax+1', '\\alpha x^10', '\\alpha x^2.5y', '\\alpha x^-1+y',
  '\\alpha x^{ab}', '\\alpha a^2+b^2', '\\alpha (x_1)W_2',
  // the equations of "Attention Is All You Need"
  '\\mathrm{Attention}(Q, K, V) = \\mathrm{softmax}\\left(\\frac{QK^T}{\\sqrt{d_k}}\\right)V',
  '\\mathrm{MultiHead}(Q, K, V) = \\mathrm{Concat}(\\mathrm{head}_1, \\ldots, \\mathrm{head}_h)W^O',
  '\\mathrm{head}_i = \\mathrm{Attention}(QW_i^Q, KW_i^K, VW_i^V)',
  '\\mathrm{FFN}(x) = \\max(0, xW_1 + b_1)W_2 + b_2',
  'PE_{(pos,2i)} = \\sin(pos/10000^{2i/d_{model}})',
  'PE_{(pos,2i+1)} = \\cos(pos/10000^{2i/d_{model}})',
  'lrate = d_{model}^{-0.5} \\cdot \\min(step\\_num^{-0.5}, step\\_num \\cdot warmup\\_steps^{-1.5})',
  // every construct, and the places where bare runs could fuse
  '\\frac{a}{b}', '\\sqrt{x}', '\\sqrt[n]{x+1}', 'x^{2}', 'x_{i}', 'x_i^2', '\\sum_i x', '\\sum x', '\\sum_{i=1}^n',
  '\\prod_{k=1}^{m} k', '\\int_0^1 f', '\\int x', '{a}{b}', '\\mathbf{x}y', '{x}{y}^2', '{\\sum_i^n} x', '\\alpha x', '\\alpha\\beta',
  '\\sin x', '\\lim_{x} f', '\\operatorname{foo} x', '\\text{hi there} x', 'a + b^2', 'a_1 + b_2', 'x^2 + y^2',
  'αx', 'a≤b', '50\\%', '50%', '\\% 5', 'a\\_b', 'x\\{1\\}', '\\$5', '\\# 1', 'a\\&b', '\\left[ x \\right]',
  '\\left\\{ x \\right\\}', '\\left. x \\right|', '\\left| x \\right|', '|x|', '\\left( a \\right)^2', 'f(x)^2',
  '\\frac{\\frac{a}{b}}{\\frac{c}{d}}', '\\sqrt{\\sqrt{x}}', 'x^{y^{z}}', '\\sum_{i}^{n}{a b}', '\\int \\frac{1}{x}',
  '\\begin{matrix} \\frac{a}{b} & x^2 \\\\ \\alpha & \\sqrt{y} \\end{matrix}',
];

describe('ommlToLatex', () => {
  it('has at least 30 formulas in the list', () => {
    expect(FORMULAS.length).toBeGreaterThanOrEqual(30);
  });

  it.each(FORMULAS)('round trip is stable: %s', (latex) => {
    const { latex2, same } = roundTrip(latex);
    expect(latex2).not.toBeNull();
    expect(same).toBe(true);
  });

  it('writes the common constructs as readable LaTeX', () => {
    const l = (s: string) => ommlToLatex(latexToOmml(doc, s));
    expect(l('\\frac{a}{b}')).toBe('\\frac{a}{b}');
    expect(l('\\alpha x')).toBe('\\alpha x');
    expect(l('\\sin x')).toBe('\\sin x');
    expect(l('\\sqrt[3]{x}')).toBe('\\sqrt[3]{x}');
    expect(l('x_i^2')).toBe('x_i^2');
    expect(l('x_{ab}^{2+1}')).toBe('x_{ab}^{2+1}');
    expect(l('\\text{hi there}')).toBe('\\text{hi there}');
  });

  it('writes plain-text math as typed', () => {
    const plain = (text: string) => {
      const r = doc.createElement('m:r');
      const t = doc.createElement('m:t');
      t.appendChild(doc.createTextNode(text));
      r.appendChild(t);
      const m = doc.createElement('m:oMath');
      m.appendChild(r);
      return ommlToLatex(m);
    };
    expect(plain('E=mc^2')).toBe('E=mc^2');
    expect(plain('a_1+b_2')).toBe('a_1+b_2');
    expect(plain('a&b')).toBe('a\\&b');
    expect(plain('^')).toBe('\\^');
  });

  it('returns null for elements the converter never emits', () => {
    const parse = (xml: string) => new DOMParser().parseFromString(`<m:oMath ${NS}>${xml}</m:oMath>`, 'text/xml').documentElement as unknown as Element;
    const cases = [
      '<m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>',
      '<m:bar><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar>',
      '<m:eqArr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:eqArr>',
      '<m:limLow><m:e><m:r><m:t>x</m:t></m:r></m:e><m:lim><m:r><m:t>y</m:t></m:r></m:lim></m:limLow>',
      '<m:r><w:rPr/><m:t>x</m:t></m:r>',
      '<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>',
      '<m:r><m:t>a b</m:t></m:r>',
      '<m:r><m:t>a]</m:t></m:r>',
      '<m:r><m:t xml:space="preserve">a{b</m:t></m:r>',
      '<m:d><m:dPr><m:begChr m:val="⟨"/><m:endChr m:val="⟩"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>',
      '<m:f><m:num/><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>',
      '',
    ];
    for (const xml of cases) expect(ommlToLatex(parse(xml))).toBeNull();
    expect(ommlToLatex(null as never)).toBeNull();
    expect(ommlToLatex(parse('<m:r><m:t>x</m:t></m:r>').firstChild as Element)).toBeNull();
  });

  it('never throws on deeply nested input', () => {
    let xml = '<m:r><m:t>x</m:t></m:r>';
    for (let i = 0; i < 200; i++) xml = `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${xml}</m:e></m:rad>`;
    const el = new DOMParser().parseFromString(`<m:oMath ${NS}>${xml}</m:oMath>`, 'text/xml').documentElement as unknown as Element;
    expect(() => ommlToLatex(el)).not.toThrow();
    expect(ommlToLatex(el)).toBeNull();
  });
});

// A seeded generator of formulas: every formula the converter accepts must round-trip.
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const ATOMS = ['x', 'y', 'ab', '12', '3.5', '+', '=', '-', 'a+b', '(x)', 'x,y', 'α', 'a≤b', '50%', '[', 'f(x)', '\\alpha', '\\beta', '\\le',
  '\\ldots', '\\infty', '\\sin', '\\lim', '\\operatorname{foo}', '\\text{a b}', '\\%', '\\_', '\\{', '\\}', '\\$', '\\#', '\\&', '\\,', '\\quad'];
const DELIMS: Array<[string, string]> = [['(', ')'], ['[', ']'], ['\\{', '\\}'], ['|', '|'], ['.', ')'], ['(', '.']];
function gen(r: () => number, depth: number): string {
  const pick = <T,>(xs: T[]): T => xs[Math.floor(r() * xs.length)];
  const seq = (n: number): string => Array.from({ length: n }, () => gen(r, depth + 1)).join(r() < 0.5 ? ' ' : '');
  if (depth >= 3 || r() < 0.35) return pick(ATOMS);
  switch (Math.floor(r() * 12)) {
    case 0: return `\\frac{${seq(1 + Math.floor(r() * 2))}}{${seq(1 + Math.floor(r() * 2))}}`;
    case 1: return `\\sqrt{${seq(1 + Math.floor(r() * 2))}}`;
    case 2: return `\\sqrt[${pick(['n', '3', 'a+b'])}]{${seq(1)}}`;
    case 3: return `${pick(['x', '{' + seq(2) + '}', '\\alpha', 'ab', '\\sin'])}^${r() < 0.5 ? '2' : '{' + seq(2) + '}'}`;
    case 4: return `${pick(['x', '{' + seq(2) + '}', '\\alpha', 'ab', '\\lim'])}_${r() < 0.5 ? 'i' : '{' + seq(2) + '}'}`;
    case 5: return `${pick(['x', 'a', '{' + seq(1) + '}'])}_{${seq(1)}}^{${seq(1)}}`;
    case 6: return `${pick(['\\sum', '\\prod', '\\int'])}${pick(['', '_{' + seq(1) + '}', '^{' + seq(1) + '}', '_{' + seq(1) + '}^{' + seq(1) + '}', '_i^n'])}${pick(['', ' x', ' ' + seq(1), '{' + seq(2) + '}'])}`;
    case 7: { const [a, b] = pick(DELIMS); return `\\left${a} ${seq(1 + Math.floor(r() * 2))} \\right${b}`; }
    case 8: return `|${seq(1)}|`;
    case 9: return `\\begin{matrix} ${seq(1)} & ${seq(1)} \\\\ ${seq(1)} & ${seq(1)} \\end{matrix}`;
    case 10: return `{${seq(1 + Math.floor(r() * 2))}}`;
    default: return `\\mathbf{${seq(1)}}`;
  }
}

describe('ommlToLatex on generated formulas', () => {
  it('every formula the converter accepts round-trips to the same OMML', () => {
    const r = rng(20261024);
    const failures: string[] = [];
    let accepted = 0;
    for (let i = 0; i < 4000; i++) {
      const count = 1 + Math.floor(r() * 4);
      const latex = Array.from({ length: count }, () => gen(r, 0)).join(r() < 0.5 ? ' ' : '');
      let omml1: Element;
      try { omml1 = latexToOmml(doc, latex); } catch { continue; }
      accepted++;
      const latex2 = ommlToLatex(omml1);
      let ok = latex2 !== null;
      if (ok) { try { ok = ser(omml1) === ser(latexToOmml(doc, latex2 as string)); } catch { ok = false; } }
      if (!ok) failures.push(`${latex}  =>  ${latex2}`);
    }
    expect(accepted).toBeGreaterThan(1500);
    expect(failures.slice(0, 15)).toEqual([]);
  }, 120000);
});
