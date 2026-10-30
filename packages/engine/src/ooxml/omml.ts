/**
 * Stage E4 — LaTeX subset → OMML converter (ECMA-376 §22.1 Office Math).
 *
 * Tokenizer (commands, grouping braces/brackets, super/subscripts, alignment
 * tabs, row breaks, literals) + recursive DOM builder with a BOUND depth
 * (OMML_MAX_DEPTH = 6 — anything deeper fails loudly, never silently).
 *
 * Structural coverage:
 *  - m:f (m:num / m:den)                    \frac
 *  - m:rad (m:radPr/m:degHide + m:deg + m:e) \sqrt, \sqrt[n]{}
 *  - m:sSup / m:sSub / m:sSubSup             ^ / _
 *  - m:nary (m:chr + m:limLoc + sub/sup/e)   \sum \prod \int
 *  - m:d (m:begChr/m:endChr + m:e)           \left..\right, |..|
 *  - m:m (m:mr rows of m:e cells)            matrix-family \begin{...} envs
 *    — bare m:m per ECMA-376 §22.1, NOT delimiter-wrapped (pmatrix-style
 *    delimiters are not preserved; documented, not silent: the env name is
 *    part of the converted structure only as rows/cells).
 *  - m:r runs, m:rPr/m:lit roman runs for named functions.
 *
 * Children are emitted in ECMA schema order in every branch
 * (f: num,den · rad: radPr,deg,e · sSubSup: e,sub,sup · nary: naryPr,sub,sup,e
 *  with naryPr: chr,limLoc · d: dPr,e with dPr: begChr,endChr · m: mr+
 *  with mr: e+). Anything unrecognized fails loudly via E_VALIDATION naming
 *  the construct + raw position + remedy — never silently dropped.
 *
 * Supersedes the E1 mini-converter (operations/equation.ts): strict superset
 * of its constructs (every E1-mapped macro maps identically here) plus
 * matrix environments, a larger symbol table, \det, and depth 4 → 6.
 */

import { createMathElement } from './ns';

/** E4 bound: structural nesting deeper than this fails loudly. */
export const OMML_MAX_DEPTH = 6;

const REMEDY =
  'simplify the expression, split it into smaller equations, or insert pre-built OMML instead of LaTeX.';

export function ommlFail(construct: string, pos: number, detail: string): never {
  const err = new Error(
    `E_VALIDATION: equation converter cannot handle ${construct} at position ${pos} (${detail}). Remedy: ${REMEDY}`
  );
  (err as unknown as Record<string, unknown>).code = 'E_VALIDATION';
  (err as unknown as Record<string, unknown>).retryable = false;
  throw err;
}

// ── Symbol table (strict superset of E1's) ──────────────────────────────────
// E1 codepoints preserved byte-identically (epsilon U+03F5, theta U+03B8, …).

const SYMBOLS: Record<string, string> = {
  // E1 set (unchanged):
  nabla: '∇', cdot: '⋅', pm: '±', mp: '∓', times: '×', div: '÷',
  le: '≤', ge: '≥', neq: '≠', approx: '≈', equiv: '≡', infty: '∞',
  partial: '∂', in: '∈', subset: '⊂',
  alpha: 'α', beta: 'β', gamma: 'γ', pi: 'π', mu: 'μ',
  epsilon: 'ϵ', theta: 'θ', sigma: 'σ', omega: 'ω',
  to: '→', rightarrow: '→',
  // E4: full lowercase Greek:
  delta: 'δ', zeta: 'ζ', eta: 'η', iota: 'ι', kappa: 'κ',
  lambda: 'λ', nu: 'ν', xi: 'ξ', omicron: 'ο', rho: 'ρ',
  tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ', psi: 'ψ',
  varepsilon: 'ε', vartheta: 'ϑ', varpi: 'ϖ', varrho: 'ϱ',
  varsigma: 'ς', varphi: 'ϕ', ell: 'ℓ',
  // E4: uppercase Greek:
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ',
  Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  // E4: arrows:
  leftarrow: '←', Leftarrow: '⇐', Rightarrow: '⇒',
  leftrightarrow: '↔', Leftrightarrow: '⇔',
  uparrow: '↑', downarrow: '↓', Uparrow: '⇑', Downarrow: '⇓',
  longleftarrow: '⟵', longrightarrow: '⟶', mapsto: '↦',
  // E4: relations:
  leq: '≤', geq: '≥', ll: '≪', gg: '≫', sim: '∼', simeq: '≃',
  cong: '≅', propto: '∝', perp: '⊥', mid: '∣', parallel: '∥',
  asymp: '≍', doteq: '≐',
  // E4: sets & logic:
  forall: '∀', exists: '∃', ni: '∋', notin: '∉',
  subseteq: '⊆', supset: '⊃', supseteq: '⊇',
  cup: '∪', cap: '∩', setminus: '∖', emptyset: '∅',
  neg: '¬', lnot: '¬', wedge: '∧', land: '∧', vee: '∨', lor: '∨',
  oplus: '⊕', ominus: '⊖', otimes: '⊗', oslash: '⊘',
  circ: '∘', bullet: '∙', ast: '∗',
  // E4: misc:
  deg: '°', prime: '′', ldots: '…', dots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱',
  aleph: 'ℵ', hbar: 'ℏ',
};

// Named functions render as roman ("literal") runs per ECMA-376 §22.1.
const FUNCS = new Set([
  'sin', 'cos', 'tan', 'lim', 'ln', 'log', 'det',
  'max', 'min', 'sup', 'inf', 'exp', 'arg', 'gcd', 'lg',
  'cot', 'sec', 'csc', 'sinh', 'cosh', 'tanh',
  'arcsin', 'arccos', 'arctan', 'limsup', 'liminf', 'dim', 'ker',
]);
// Formatting wrappers are stripped to their (recursively converted) content.
/** Accent commands -> combining character for m:accPr/m:chr (ECMA-376 §22.1.2.1). */
const ACCENTS: Record<string, string> = {
  dot: '\u0307',
  ddot: '\u0308',
  hat: '\u0302',
  bar: '\u0304',
  vec: '\u20D7',
  tilde: '\u0303',
};
const WRAPPERS = new Set(['mathbf', 'mathrm', 'mathit', 'boldsymbol', 'mathsf']);
// N-ary operators: sum-like take undOvr limits, integrals take subSup.
const NARY: Record<string, { chr: string; limLoc: string }> = {
  sum: { chr: '∑', limLoc: 'undOvr' },
  prod: { chr: '∏', limLoc: 'undOvr' },
  int: { chr: '∫', limLoc: 'subSup' },
};
// Matrix environments → bare m:m (NOT delimiter-wrapped, per E4 spec).
const MATRIX_ENVS = new Set([
  'matrix', 'pmatrix', 'bmatrix', 'Bmatrix', 'vmatrix', 'Vmatrix',
  'smallmatrix', 'array',
]);

// ── Tokenizer ───────────────────────────────────────────────────────────────

type Tok =
  | { t: 'lit'; text: string; pos: number; gap: boolean }
  | { t: 'cmd'; name: string; pos: number }
  | { t: 'lbrace'; pos: number }
  | { t: 'rbrace'; pos: number }
  | { t: 'lbracket'; pos: number; gap: boolean }
  | { t: 'rbracket'; pos: number }
  | { t: 'sup'; pos: number }
  | { t: 'sub'; pos: number }
  | { t: 'amp'; pos: number }
  | { t: 'bar'; pos: number }
  | { t: 'begin'; env: string; colspec: string | null; pos: number }
  | { t: 'end'; env: string | null; pos: number }
  | { t: 'textlit'; text: string; pos: number; roman?: boolean };

function isAlpha(c: string): boolean {
  return /^[a-zA-Z]$/.test(c);
}

/** Raw `{env}` read at i (s[i] === '{'). Returns [inner, after] or null. */
function readBraceRaw(s: string, i: number): [string, number] | null {
  if (s[i] !== '{') return null;
  const close = s.indexOf('}', i);
  if (close < 0) return null;
  return [s.slice(i + 1, close), close + 1];
}

function tokenize(s: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const n = s.length;
  // gap: whitespace was skipped since the previous token. Content runs merge
  // across whitespace (math-mode skipping), but operand WORDS do not (E1:
  // `a^2 + b` takes `2`, not `2+b`, as the superscript) — lit/lbracket carry
  // the flag so parseOperand can reproduce the boundary.
  let gap = false;
  const markGap = (): boolean => {
    const g = gap;
    gap = false;
    return g;
  };
  while (i < n) {
    const c = s[i];
    if (/\s/.test(c)) { i++; gap = true; continue; }
    if (c === '{') { toks.push({ t: 'lbrace', pos: i }); gap = false; i++; continue; }
    if (c === '}') { toks.push({ t: 'rbrace', pos: i }); gap = false; i++; continue; }
    if (c === '[') { toks.push({ t: 'lbracket', pos: i, gap: markGap() }); i++; continue; }
    if (c === ']') { toks.push({ t: 'rbracket', pos: i }); gap = false; i++; continue; }
    if (c === '^') { toks.push({ t: 'sup', pos: i }); gap = false; i++; continue; }
    if (c === '_') { toks.push({ t: 'sub', pos: i }); gap = false; i++; continue; }
    if (c === '&') { toks.push({ t: 'amp', pos: i }); gap = false; i++; continue; }
    if (c === '|') { toks.push({ t: 'bar', pos: i }); gap = false; i++; continue; }
    if (c === '\\') {
      const bs = i;
      i++;
      if (i >= n) ommlFail('trailing backslash', bs, `a lone '\\' at the end is not convertible`);
      let name: string;
      if (isAlpha(s[i])) {
        const m = /^[a-zA-Z]+/.exec(s.slice(i));
        name = (m as RegExpExecArray)[0];
        i += name.length;
      } else {
        name = s[i];
        i++;
      }
      if (name === 'begin' || name === 'end') {
        let j = i;
        while (j < n && /\s/.test(s[j])) j++;
        const raw = readBraceRaw(s, j);
        if (!raw) {
          // Malformed: no {env} — parser fails naming the bare command (E1-identical).
          toks.push({ t: 'cmd', name, pos: bs });
          continue;
        }
        const [env, after] = raw;
        i = after;
        if (name === 'end') {
          toks.push({ t: 'end', env, pos: bs });
          continue;
        }
        let colspec: string | null = null;
        if (env === 'array') {
          let k = i;
          while (k < n && /\s/.test(s[k])) k++;
          const crav = readBraceRaw(s, k);
          if (!crav) {
            ommlFail(`'\\begin{array}'`, bs, 'expected {columns} specifier (e.g. {cc})');
          }
          colspec = crav[0];
          if (!/^[lcr\s]+$/.test(colspec)) {
            ommlFail(
              `'\\begin{array}'`, bs,
              `invalid column specifier {${colspec}} (only l, c, r are supported; vertical rules | are not supported)`
            );
          }
          i = crav[1];
        }
        toks.push({ t: 'begin', env, colspec, pos: bs });
        continue;
      }
      if (name === 'text' || name === 'mbox' || name === 'operatorname') {
        // Space-preserving raw group (E1-identical semantics + messages).
        let j = i;
        while (j < n && /\s/.test(s[j])) j++;
        if (s[j] !== '{') {
          toks.push({ t: 'cmd', name, pos: bs });
          continue; // parser fails `'\\text'` + `expected {argument}` at the next token.
        }
        j++; // consume '{'
        let out = '';
        let inner = 0;
        while (j < n) {
          const d = s[j];
          if (d === '{') { inner++; j++; continue; }
          if (d === '}') {
            if (inner === 0) break;
            inner--;
            j++;
            continue;
          }
          out += d;
          j++;
        }
        if (j >= n) ommlFail(`'\\${name}'`, bs, `missing closing '}'`);
        i = j + 1; // consume '}'
        if (out === '') ommlFail(`'empty \\${name}{}'`, bs, 'text groups must contain literal content');
        if (name === 'operatorname' && !/^[A-Za-z]+$/.test(out)) {
          ommlFail(`'\\operatorname'`, bs, 'only letters are supported inside \\operatorname{...}');
        }
        toks.push({ t: 'textlit', text: out, pos: bs, roman: name === 'operatorname' });
        continue;
      }
      toks.push({ t: 'cmd', name, pos: bs });
      continue;
    }
    // Literal run: everything not otherwise handled, minus whitespace.
    // (All brackets/commands are tokenized above, so a run is always non-empty.)
    let text = '';
    const start = i;
    while (i < n && !/[{}[\]\\^_&|\s]/.test(s[i])) {
      text += s[i++];
    }
    toks.push({ t: 'lit', text, pos: start, gap: markGap() });
  }
  return toks;
}

// ── Recursive DOM builder ───────────────────────────────────────────────────

interface Ctx {
  s: string;
  toks: Tok[];
  i: number;
}

interface Stop {
  brace: boolean;
  right: boolean;
  bar: boolean;
  bracket: boolean;
  cell: boolean; // matrix cell: stop at & | \\ | \end
}

type StopReason = 'end' | 'brace' | 'right' | 'bar' | 'bracket' | 'amp' | 'row' | 'endenv';

const NO_STOP: Stop = {
  brace: false, right: false, bar: false, bracket: false, cell: false,
};

function nest(depth: number, pos: number): number {
  const next = depth + 1;
  if (next > OMML_MAX_DEPTH) {
    ommlFail('deeply nested expression', pos, `nesting exceeds ${OMML_MAX_DEPTH} levels`);
  }
  return next;
}

function peek(ctx: Ctx): Tok | null {
  return ctx.i < ctx.toks.length ? ctx.toks[ctx.i] : null;
}

function runNode(doc: Document, text: string): Element {
  const r = createMathElement(doc, 'm:r');
  const t = createMathElement(doc, 'm:t');
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

function litRun(doc: Document, text: string): Element {
  const r = createMathElement(doc, 'm:r');
  const rPr = createMathElement(doc, 'm:rPr');
  rPr.appendChild(createMathElement(doc, 'm:lit'));
  r.appendChild(rPr);
  const t = createMathElement(doc, 'm:t');
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

function textlitRun(doc: Document, text: string): Element {
  const r = createMathElement(doc, 'm:r');
  const t = createMathElement(doc, 'm:t');
  t.setAttribute('xml:space', 'preserve');
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

function wrapSubSup(doc: Document, base: Element[], sub: Element[] | null, sup: Element[] | null): Element {
  const tag = sub && sup ? 'm:sSubSup' : sub ? 'm:sSub' : 'm:sSup';
  const el = createMathElement(doc, tag);
  const e = createMathElement(doc, 'm:e');
  for (const nd of base) e.appendChild(nd);
  el.appendChild(e);
  if (sub) {
    const s = createMathElement(doc, 'm:sub');
    for (const nd of sub) s.appendChild(nd);
    el.appendChild(s);
  }
  if (sup) {
    const s = createMathElement(doc, 'm:sup');
    for (const nd of sup) s.appendChild(nd);
    el.appendChild(s);
  }
  return el;
}

function fragKids(frag: DocumentFragment): Element[] {
  return Array.from(frag.childNodes) as Element[];
}

function tokLabel(t: Tok): string {
  switch (t.t) {
    case 'lit': return `'${t.text}'`;
    case 'cmd': return `'\\${t.name}'`;
    case 'lbrace': return `'{'`;
    case 'rbrace': return `'}'`;
    case 'lbracket': return `'['`;
    case 'rbracket': return `']'`;
    case 'sup': return `'^'`;
    case 'sub': return `'_'`;
    case 'amp': return `'&'`;
    case 'bar': return `'|'`;
    case 'begin': return `'\\begin{${t.env}}'`;
    case 'end': return t.env === null ? `'\\end'` : `'\\end{${t.env}}'`;
    case 'textlit': return 'text';
  }
}

function eofPos(ctx: Ctx): number {
  return ctx.s.length;
}

function nextPos(ctx: Ctx): number {
  const t = peek(ctx);
  return t ? t.pos : eofPos(ctx);
}

function parseContent(
  doc: Document, ctx: Ctx, parent: Element | DocumentFragment, depth: number, stop: Stop
): StopReason {
  let buf = '';
  const flush = (): void => {
    if (buf !== '') {
      parent.appendChild(runNode(doc, buf));
      buf = '';
    }
  };
  for (;;) {
    const tok = peek(ctx);
    if (!tok) { flush(); return 'end'; }
    switch (tok.t) {
      case 'rbrace':
        if (stop.brace) { flush(); return 'brace'; }
        ommlFail(`unbalanced brace '}'`, tok.pos, `unexpected '}' with no matching '{'`);
        break;
      case 'rbracket':
        if (stop.bracket) { flush(); return 'bracket'; }
        ommlFail(`unbalanced bracket ']'`, tok.pos, `unexpected ']' with no matching '[' (did you mean '\\sqrt[degree]{...}'?)`);
        break;
      case 'bar':
        if (stop.bar) { flush(); return 'bar'; }
        flush();
        ctx.i++;
        appendSuffixed(doc, ctx, [parseBar(doc, ctx, depth)], depth, parent);
        continue;
      case 'cmd':
        if (tok.name === 'right' && stop.right) { flush(); return 'right'; }
        if (tok.name === '\\' && stop.cell) { flush(); return 'row'; }
        flush();
        ctx.i++;
        appendSuffixed(doc, ctx, parseCommand(doc, ctx, depth, tok), depth, parent);
        continue;
      case 'begin':
      case 'lbrace':
      case 'textlit': {
        flush();
        const base = parseBase(doc, ctx, depth);
        appendSuffixed(doc, ctx, base, depth, parent);
        continue;
      }
      case 'end':
        if (stop.cell) { flush(); return 'endenv'; }
        ctx.i++;
        ommlFail(
          `unsupported '\\end{${tok.env ?? ''}}' environment`, tok.pos,
          'matrices, cell alignment and multi-line environments are not supported outside a matching \\begin'
        );
        break;
      case 'amp':
        if (stop.cell) { flush(); return 'amp'; }
        ommlFail('alignment tab (&)', tok.pos, 'matrices and cell alignment (&, \\\\, \\begin{...}) are not supported here');
        break;
      case 'sup':
      case 'sub': {
        // Plain-text base (e.g. the `b` in `b^2`): wrap the buffered run.
        const mark = tok.t === 'sup' ? '^' : '_';
        if (buf === '') {
          ommlFail(`stray '${mark}'`, tok.pos, `'${mark}' must follow a base expression (e.g. x${mark}{2})`);
        }
        // A leading relation (`=`, `<`, `>`) has no left operand inside a
        // base and renders as a broken glyph in LibreOffice (PDF export), so it
        // stays a run of its own in front of the script. Other characters keep
        // the E1 rule (the whole buffered run is the base).
        const lead = /^[=<>]+(?=.)/s.exec(buf);
        if (lead) {
          parent.appendChild(runNode(doc, lead[0]));
          buf = buf.slice(lead[0].length);
        }
        const base = [runNode(doc, buf)];
        buf = '';
        appendSuffixed(doc, ctx, base, depth, parent);
        continue;
      }
      case 'lbracket':
        // E1: '[' is literal text in content position.
        buf += '[';
        ctx.i++;
        continue;
      case 'lit':
        buf += tok.text;
        ctx.i++;
        continue;
    }
  }
}

function appendSuffixed(
  doc: Document, ctx: Ctx, base: Element[], depth: number, parent: Element | DocumentFragment
): void {
  for (const nd of parseSuffixes(doc, ctx, base, depth)) parent.appendChild(nd);
}

/** One atom: {group} | \command | |bar| | \begin-env | textlit. Never consumes ^/_/&/}/stops. */
function parseBase(doc: Document, ctx: Ctx, depth: number): Element[] {
  const tok = peek(ctx);
  if (!tok) return ommlFail('dangling argument', eofPos(ctx), 'expected an argument but reached the end of the expression');
  switch (tok.t) {
    case 'lbrace':
      return parseGroup(doc, ctx, depth);
    case 'bar':
      ctx.i++;
      return [parseBar(doc, ctx, depth)];
    case 'cmd': {
      const cmd = tok;
      ctx.i++;
      return parseCommand(doc, ctx, depth, cmd);
    }
    case 'begin': {
      const b = tok;
      ctx.i++;
      return [parseBeginEnv(doc, ctx, depth, b)];
    }
    case 'textlit':
      ctx.i++;
      return [tok.roman ? litRun(doc, tok.text) : textlitRun(doc, tok.text)];
    default:
      return ommlFail(`unexpected ${tokLabel(tok)}`, tok.pos, 'internal converter error');
  }
}

function parseGroup(doc: Document, ctx: Ctx, depth: number): Element[] {
  const open = peek(ctx) as { pos: number };
  ctx.i++; // consume '{'
  const nested = nest(depth, open.pos);
  const frag = doc.createDocumentFragment();
  const r = parseContent(doc, ctx, frag, nested, { ...NO_STOP, brace: true });
  if (r !== 'brace') ommlFail(`unbalanced brace '{'`, open.pos, `missing closing '}'`);
  ctx.i++; // consume '}'
  const kids = fragKids(frag);
  if (kids.length === 0) ommlFail('empty group {}', open.pos, 'groups must contain convertible content');
  return kids;
}

/** '{...}' argument required by structural commands. */
function parseBraceArg(doc: Document, ctx: Ctx, depth: number, cmd: string, what: string): Element[] {
  const tok = peek(ctx);
  if (!tok || tok.t !== 'lbrace') ommlFail(`'${cmd}'`, nextPos(ctx), `expected ${what}`);
  return parseGroup(doc, ctx, depth);
}

/**
 * Length of the plain-text operand that ^ or _ takes from the start of a literal token.
 * Rule: an optional sign, then either one balanced (...) group, or a whole number (with one
 * decimal part), or a single letter, else a single character. So `2+b` gives `2`,
 * `10+y` gives `10`, `1)W` gives `1`, `n,m` gives `n`, `ij` gives `i`, `2y` gives `2`,
 * `ax` gives `a`, `-1` gives `-1`, `(n)+1` gives `(n)`. A number keeps its whole digit
 * run (and one decimal part) because `x^10` is overwhelmingly meant as ten; this is the
 * only intended difference from strict TeX.
 */
function operandPrefixLength(text: string): number {
  let i = 0;
  if (text[i] === '+' || text[i] === '-') i++;
  if (text[i] === '(') {
    let depth = 0;
    for (let j = i; j < text.length; j++) {
      if (text[j] === '(') depth++;
      else if (text[j] === ')') {
        depth--;
        if (depth === 0) return j + 1;
      }
    }
    return i + 1;
  }
  // TeX takes one token: a single letter, or (as a practical reading of x^10) a whole number.
  const m = /^(?:[0-9]+(?:\.[0-9]+)?|[A-Za-z])/.exec(text.slice(i));
  if (m) return i + m[0].length;
  return i > 0 ? i : 1;
}

/**
 * One operand: a base atom or a plain-text word. Suffixes are taken only when
 * takeSuffixes (the n-ary integrand); limit/argument position takes exactly
 * one operand (TeX-consistent: `x^2^3` stays a loud duplicate).
 */
function parseOperand(doc: Document, ctx: Ctx, depth: number, takeSuffixes: boolean): Element[] {
  const tok = peek(ctx);
  if (!tok) {
    ommlFail('dangling argument', eofPos(ctx), 'expected an argument but reached the end of the expression');
  }
  const t = tok as Tok;
  if (t.t === 'lbrace' || t.t === 'cmd' || t.t === 'bar' || t.t === 'begin' || t.t === 'textlit') {
    const base = parseBase(doc, ctx, depth);
    return takeSuffixes ? parseSuffixes(doc, ctx, base, depth) : base;
  }
  if (
    t.t === 'rbrace' || t.t === 'rbracket' || t.t === 'sup' || t.t === 'sub' ||
    t.t === 'amp' || t.t === 'end'
  ) {
    return ommlFail(`unexpected ${tokLabel(t)}`, t.pos, 'expected an argument here');
  }
  if (!takeSuffixes) {
    // ^ and _ operands (and limit arguments) take only the operand prefix of the first
    // literal token; the rest of the token goes back for the caller to read as content.
    const first = peek(ctx);
    if (first && first.t === 'lit') {
      const cut = operandPrefixLength(first.text);
      if (cut < first.text.length) {
        toksReplaceHead(ctx, { t: 'lit', text: first.text.slice(cut), pos: first.pos + cut, gap: false });
        return [runNode(doc, first.text.slice(0, cut))];
      }
    }
  }
  // Plain-text word (E1: maximal whitespace-free run; '[' allowed inside,
  // ']' and whitespace end it). The first token's gap is irrelevant; later
  // tokens with gap=true belonged to a separate source run — stop there.
  let w = '';
  let started = false;
  for (;;) {
    const u = peek(ctx);
    if (!u) break;
    if (u.t === 'lit') {
      if (started && u.gap) break;
      w += u.text;
      ctx.i++;
      started = true;
    } else if (u.t === 'lbracket') {
      if (started && u.gap) break;
      w += '[';
      ctx.i++;
      started = true;
    } else break;
  }
  const base = [runNode(doc, w)];
  return takeSuffixes ? parseSuffixes(doc, ctx, base, depth) : base;
}

/** Attaches trailing ^/_ limits: wraps in m:sSup/m:sSub/m:sSubSup. */
function parseSuffixes(doc: Document, ctx: Ctx, base: Element[], depth: number): Element[] {
  let sub: Element[] | null = null;
  let sup: Element[] | null = null;
  for (;;) {
    const tok = peek(ctx);
    if (!tok || (tok.t !== 'sup' && tok.t !== 'sub')) break;
    const isSub = tok.t === 'sub';
    if (base.length === 0) ommlFail(`stray '${isSub ? '_' : '^'}'`, tok.pos, `'${isSub ? '_' : '^'}' must follow a base expression`);
    if (base.length === 1 && (base[0] as Element).tagName === 'm:nary') {
      ommlFail(`duplicate ${isSub ? 'subscript' : 'superscript'} limit`, tok.pos, 'n-ary operators take at most one lower and one upper limit');
    }
    if (isSub && sub) ommlFail('duplicate subscript', tok.pos, 'a base takes at most one subscript (use groups for the rest)');
    if (!isSub && sup) ommlFail('duplicate superscript', tok.pos, 'a base takes at most one superscript (use groups for the rest)');
    ctx.i++;
    const arg = parseOperand(doc, ctx, nest(depth, tok.pos), false);
    if (isSub) sub = arg; else sup = arg;
  }
  if (!sub && !sup) return base;
  return [wrapSubSup(doc, base, sub, sup)];
}

function parseBar(doc: Document, ctx: Ctx, depth: number): Element {
  const open = (ctx.toks[ctx.i - 1] as Tok).pos;
  const nested = nest(depth, open);
  const frag = doc.createDocumentFragment();
  const r = parseContent(doc, ctx, frag, nested, { ...NO_STOP, bar: true });
  if (r !== 'bar') ommlFail(`unmatched '|' delimiter`, open, `missing closing '|'`);
  ctx.i++; // consume closing '|'
  const kids = fragKids(frag);
  if (kids.length === 0) ommlFail(`empty '|..|' delimiter`, open, 'delimited content must be convertible');
  const d = createMathElement(doc, 'm:d');
  const dPr = createMathElement(doc, 'm:dPr');
  const beg = createMathElement(doc, 'm:begChr');
  beg.setAttribute('m:val', '|');
  const end = createMathElement(doc, 'm:endChr');
  end.setAttribute('m:val', '|');
  dPr.appendChild(beg);
  dPr.appendChild(end);
  d.appendChild(dPr);
  const e = createMathElement(doc, 'm:e');
  for (const nd of kids) e.appendChild(nd);
  d.appendChild(e);
  return d;
}

const LEFT_DELIMS = '()[]{}|.';

function parseLeft(doc: Document, ctx: Ctx, depth: number, bs: number): Element {
  const takeDelim = (which: 'left' | 'right'): string => {
    const tok = peek(ctx);
    if (!tok) ommlFail(`'\\${which}'`, eofPos(ctx), 'expected a delimiter ( ( [ { | . )');
    const t = tok as Tok;
    if (t.t === 'lbrace') { ctx.i++; return '{'; }
    if (t.t === 'rbrace') { ctx.i++; return '}'; }
    if (t.t === 'lbracket') { ctx.i++; return '['; }
    if (t.t === 'rbracket') { ctx.i++; return ']'; }
    if (t.t === 'bar') { ctx.i++; return '|'; }
    // Escaped delimiters: \{ \} \| tokenize as cmd tokens named '{' '}' '|'.
    if (t.t === 'cmd' && (t.name === '{' || t.name === '}' || t.name === '|')) {
      ctx.i++;
      return t.name;
    }
    if (t.t === 'lit' && t.text.length > 0 && LEFT_DELIMS.includes(t.text[0])) {
      const d = t.text[0];
      if (t.text.length === 1) ctx.i++;
      else toksReplaceHead(ctx, { t: 'lit', text: t.text.slice(1), pos: t.pos + 1, gap: false });
      return d;
    }
    return ommlFail(
      `'\\${which}'`, t.pos,
      which === 'left' ? 'expected a delimiter ( ( [ { | . )' : 'expected a delimiter ( ) ] } | . )'
    );
  };
  const d0 = takeDelim('left');
  const nested = nest(depth, bs);
  const frag = doc.createDocumentFragment();
  const r = parseContent(doc, ctx, frag, nested, { ...NO_STOP, right: true });
  if (r !== 'right') ommlFail(`'\\left'`, bs, `missing '\\right'`);
  ctx.i++; // consume '\right'
  const d1 = takeDelim('right');
  const kids = fragKids(frag);
  if (kids.length === 0) ommlFail(`'\\left..\\right'`, bs, 'delimited content must be convertible');
  const d = createMathElement(doc, 'm:d');
  // Always emit dPr with explicit begChr/endChr: a missing begChr/endChr
  // defaults to "(" / ")" per ECMA-376 (CT_DelimiterPr defaults), so Word
  // would render a parenthesis for an invisible '.' side. m:val="" is
  // schema-legal (optional ST_Xstring) and renders invisibly.
  const dPr = createMathElement(doc, 'm:dPr');
  const beg = createMathElement(doc, 'm:begChr');
  beg.setAttribute('m:val', d0 === '.' ? '' : d0);
  dPr.appendChild(beg);
  const end = createMathElement(doc, 'm:endChr');
  end.setAttribute('m:val', d1 === '.' ? '' : d1);
  dPr.appendChild(end);
  d.appendChild(dPr);
  const e = createMathElement(doc, 'm:e');
  for (const nd of kids) e.appendChild(nd);
  d.appendChild(e);
  return d;
}

/** Replace the head token (used to split one delimiter char off a literal run). */
function toksReplaceHead(ctx: Ctx, tok: Tok): void {
  ctx.toks[ctx.i] = tok;
}

function parseNary(doc: Document, ctx: Ctx, depth: number, name: string): Element[] {
  const info = NARY[name];
  let sub: Element[] | null = null;
  let sup: Element[] | null = null;
  for (;;) {
    const tok = peek(ctx);
    if (tok && tok.t === 'sub' && !sub) { ctx.i++; sub = parseOperand(doc, ctx, nest(depth, tok.pos), false); }
    else if (tok && tok.t === 'sup' && !sup) { ctx.i++; sup = parseOperand(doc, ctx, nest(depth, tok.pos), false); }
    else break;
  }
  const chrRun = [runNode(doc, info.chr)];
  // The integrand is the next operand; absent at end-of-input, '}' / ']' /
  // a closing '\right' (which belongs to an enclosing \left, not to us).
  const ahead = peek(ctx);
  const noE = !ahead || ahead.t === 'rbrace' || ahead.t === 'rbracket' ||
    (ahead.t === 'cmd' && ahead.name === 'right');
  const eNodes = noE ? null : parseOperand(doc, ctx, nest(depth, (ahead as Tok).pos), true);
  if (sub && sup && eNodes) {
    // Schema-valid n-ary requires naryPr + sub + sup + e, in order.
    const nary = createMathElement(doc, 'm:nary');
    const pr = createMathElement(doc, 'm:naryPr');
    const chr = createMathElement(doc, 'm:chr');
    chr.setAttribute('m:val', info.chr);
    const lim = createMathElement(doc, 'm:limLoc');
    lim.setAttribute('m:val', info.limLoc);
    pr.appendChild(chr);
    pr.appendChild(lim);
    nary.appendChild(pr);
    const subEl = createMathElement(doc, 'm:sub');
    for (const nd of sub) subEl.appendChild(nd);
    nary.appendChild(subEl);
    const supEl = createMathElement(doc, 'm:sup');
    for (const nd of sup) supEl.appendChild(nd);
    nary.appendChild(supEl);
    const eEl = createMathElement(doc, 'm:e');
    for (const nd of eNodes) eEl.appendChild(nd);
    nary.appendChild(eEl);
    return [nary];
  }
  // Fewer than (both limits + integrand): stay schema-valid by emitting the
  // operator as a plain run with sSub/sSup wrappers; a present integrand
  // follows as ordinary siblings.
  const out = sub || sup ? [wrapSubSup(doc, chrRun, sub, sup)] : chrRun;
  if (eNodes) out.push(...eNodes);
  return out;
}

function parseBeginEnv(
  doc: Document, ctx: Ctx, depth: number,
  b: { env: string; colspec: string | null; pos: number }
): Element {
  if (!MATRIX_ENVS.has(b.env)) {
    ommlFail(
      `'\\begin{${b.env}}' environment`, b.pos,
      'matrices, cell alignment and multi-line environments are not supported'
    );
  }
  return parseMatrix(doc, ctx, depth, b.env, b.pos);
}

/**
 * Matrix-family environment → bare m:m (m:mr rows of m:e cells), ECMA-376
 * §22.1 CT_M. Delimited variants (pmatrix/bmatrix/…) are NOT wrapped in m:d
 * (E4 spec): the delimiter characters are not preserved.
 */
function parseMatrix(doc: Document, ctx: Ctx, depth: number, env: string, bs: number): Element {
  const nested = nest(depth, bs);
  const rows: Element[][][] = [];
  let cells: Element[][] = [];
  for (;;) {
    const frag = doc.createDocumentFragment();
    const r = parseContent(doc, ctx, frag, nested, { ...NO_STOP, cell: true });
    const kids = fragKids(frag);
    if (r === 'end') {
      ommlFail(`'\\begin{${env}}'`, bs, `missing '\\end{${env}}'`);
    }
    if (kids.length === 0) {
      const at = nextPos(ctx);
      ommlFail('empty matrix cell', at, `every cell in '\\begin{${env}}' must contain convertible content (use 0 or \\cdot, never leave a cell empty)`);
    }
    cells.push(kids);
    if (r === 'amp') {
      ctx.i++; // consume '&'
      continue;
    }
    if (r === 'row') {
      ctx.i++; // consume '\\'
      rows.push(cells);
      cells = [];
      continue;
    }
    // r === 'endenv'
    const endTok = peek(ctx);
    ctx.i++; // consume '\end{...}'
    if (!endTok || endTok.t !== 'end' || endTok.env !== env) {
      ommlFail(
        `mismatched '\\end{${endTok && endTok.t === 'end' ? (endTok.env ?? '') : ''}}'`, nextPos(ctx),
        `expected '\\end{${env}}' to close '\\begin{${env}}'`
      );
    }
    rows.push(cells);
    break;
  }
  if (rows.length === 0) {
    ommlFail(`'\\begin{${env}}'`, bs, 'matrix environments must contain at least one row');
  }
  const m = createMathElement(doc, 'm:m');
  for (const row of rows) {
    const mr = createMathElement(doc, 'm:mr');
    for (const cell of row) {
      const e = createMathElement(doc, 'm:e');
      for (const nd of cell) e.appendChild(nd);
      mr.appendChild(e);
    }
    m.appendChild(mr);
  }
  return m;
}

function parseCommand(
  doc: Document, ctx: Ctx, depth: number, tok: { name: string; pos: number }
): Element[] {
  const bs = tok.pos;
  const name = tok.name;

  if (name === 'frac') {
    const num = parseBraceArg(doc, ctx, depth, '\\frac', '{numerator}{denominator}');
    const den = parseBraceArg(doc, ctx, depth, '\\frac', '{numerator}{denominator}');
    const f = createMathElement(doc, 'm:f');
    const numEl = createMathElement(doc, 'm:num');
    for (const nd of num) numEl.appendChild(nd);
    const denEl = createMathElement(doc, 'm:den');
    for (const nd of den) denEl.appendChild(nd);
    f.appendChild(numEl);
    f.appendChild(denEl);
    return [f];
  }
  if (name === 'sqrt') {
    let deg: Element[] | null = null;
    const ahead = peek(ctx);
    if (ahead && ahead.t === 'lbracket') {
      ctx.i++; // consume '['
      const nested = nest(depth, bs);
      const frag = doc.createDocumentFragment();
      const r = parseContent(doc, ctx, frag, nested, { ...NO_STOP, bracket: true });
      if (r !== 'bracket') ommlFail(`'\\sqrt'`, bs, `missing ']' for the optional root degree`);
      ctx.i++; // consume ']'
      deg = fragKids(frag);
      if (deg.length === 0) ommlFail(`'\\sqrt'`, bs, 'empty root degree [...] is not convertible');
    }
    const arg = parseBraceArg(doc, ctx, depth, '\\sqrt', '{radicand}');
    const rad = createMathElement(doc, 'm:rad');
    const pr = createMathElement(doc, 'm:radPr');
    const hide = createMathElement(doc, 'm:degHide');
    hide.setAttribute('m:val', deg ? '0' : '1');
    pr.appendChild(hide);
    rad.appendChild(pr);
    const degEl = createMathElement(doc, 'm:deg');
    if (deg) for (const nd of deg) degEl.appendChild(nd);
    rad.appendChild(degEl);
    const eEl = createMathElement(doc, 'm:e');
    for (const nd of arg) eEl.appendChild(nd);
    rad.appendChild(eEl);
    return [rad];
  }
  if (name === 'sum' || name === 'prod' || name === 'int') {
    return parseNary(doc, ctx, depth, name);
  }
  if (name === 'left') return [parseLeft(doc, ctx, depth, bs)];
  if (name === 'right') {
    ommlFail(`'\\right' without matching '\\left'`, bs, `'\\right' must close a '\\left' delimiter`);
  }
  if (name === 'begin' || name === 'end') {
    // Tokenizer fallback for a bare \begin / \end with no {env} (E1-identical).
    ommlFail(`unsupported '\\${name}' environment`, bs, 'matrices, cell alignment and multi-line environments are not supported');
  }
  if (name === 'hline' || name === 'cr') {
    ommlFail(`unsupported '\\${name}'`, bs, 'table rules and row breaks (matrices, alignment) are not supported');
  }
  if (Object.prototype.hasOwnProperty.call(SYMBOLS, name)) return [runNode(doc, SYMBOLS[name])];
  if (FUNCS.has(name)) return [litRun(doc, name)];
  if (name === 'text' || name === 'mbox') {
    // Tokenizer only reaches here when no {argument} followed (E1-identical).
    ommlFail(`'\\${name}'`, nextPos(ctx), `expected {argument}`);
  }
  if (Object.prototype.hasOwnProperty.call(ACCENTS, name)) {
    const arg = parseBraceArg(doc, ctx, depth, `\\${name}`, '{argument}');
    const acc = createMathElement(doc, 'm:acc');
    const accPr = createMathElement(doc, 'm:accPr');
    const chr = createMathElement(doc, 'm:chr');
    chr.setAttribute('m:val', ACCENTS[name]);
    accPr.appendChild(chr);
    acc.appendChild(accPr);
    const eEl = createMathElement(doc, 'm:e');
    for (const nd of arg) eEl.appendChild(nd);
    acc.appendChild(eEl);
    return [acc];
  }
  if (WRAPPERS.has(name)) {
    return parseBraceArg(doc, ctx, depth, `\\${name}`, '{argument}');
  }
  // Cosmetic spacing commands normalize to nothing (math spacing is implicit).
  if (
    name === ',' || name === ';' || name === ':' || name === '!' ||
    name === ' ' || name === '\n' || name === '\t' ||
    name === 'quad' || name === 'qquad'
  ) {
    return [];
  }
  if (name === '\\') {
    ommlFail(`'\\\\' line break / alignment`, bs, 'line breaks and alignment (\\\\, &, \\begin{...}) are not supported');
  }
  // Escaped literals: \& \% \# \$ \_ \^ \{ \}
  if (
    name === '%' || name === '#' || name === '$' || name === '&' ||
    name === '_' || name === '^' || name === '{' || name === '}'
  ) {
    return [runNode(doc, name)];
  }
  return ommlFail(`unsupported macro '\\${name}'`, bs, 'this LaTeX command has no OMML mapping in the E4 subset');
}

/**
 * Converts backslash-bearing latex into a structural m:oMath element.
 * Loud (E_VALIDATION) on anything unsupported — never silent.
 */
export function latexToOmml(doc: Document, latex: string): Element {
  const ctx: Ctx = { s: latex, toks: tokenize(latex), i: 0 };
  const oMath = createMathElement(doc, 'm:oMath');
  parseContent(doc, ctx, oMath, 0, NO_STOP);
  if (oMath.childNodes.length === 0) {
    ommlFail('empty equation', 0, 'the expression produced no convertible content');
  }
  return oMath;
}

// ── OMML → LaTeX ────────────────────────────────────────────────────────────
// The inverse of latexToOmml for the structures it emits. Contract: for every
// LaTeX string L that latexToOmml accepts, latexToOmml(ommlToLatex(latexToOmml(L)))
// is the same XML as latexToOmml(L). Anything the converter does not emit (Word
// accents, bars, equation arrays, control properties, unknown attributes) makes the
// whole call return null, never partial or wrong LaTeX, and the function never throws.

/** First name wins when two commands give one character (ldots before dots, le before leq). */
const SYMBOL_NAME_BY_CHAR: Record<string, string> = (() => {
  const inv: Record<string, string> = {};
  for (const [name, ch] of Object.entries(SYMBOLS)) if (!(ch in inv)) inv[ch] = name;
  return inv;
})();
const NARY_NAME_BY_CHR: Record<string, string> = { '∑': 'sum', '∏': 'prod', '∫': 'int' };
const LATEX_MAX_DEPTH = 48;

class NotConvertible extends Error {}
function notConvertible(): never {
  throw new NotConvertible();
}

/** One converted node. The flags say how it can merge with its neighbours once printed. */
interface LatexPiece {
  s: string;
  /** Starts with bare text: it would fuse with a preceding bare-text run. */
  startsWord: boolean;
  /** Ends with bare text: a following bare-text run would fuse into it. */
  endsWord: boolean;
  /** Ends in a command name such as `\alpha`: a following letter needs a space. */
  endsCmd: boolean;
  /** A sum or integral with both limits and no integrand (see latexSeq). */
  fenceIfFollowed?: boolean;
  /** Ends with a sum, product or integral sign that has no integrand yet (see latexSeq). */
  endsNary?: boolean;
}

const piece = (s: string, startsWord = false, endsWord = false, endsCmd = false): LatexPiece =>
  ({ s, startsWord, endsWord, endsCmd });

function elementKids(el: Element): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) out.push(n as Element);
    else if (n.nodeType === 8) continue;
    else if (n.nodeType === 3 && (n.nodeValue ?? '').trim() === '') continue;
    else notConvertible();
  }
  return out;
}

function onlyAttrs(el: Element, names: string[]): void {
  for (let i = 0; i < el.attributes.length; i++) {
    if (!names.includes(el.attributes[i].name)) notConvertible();
  }
}

function exactKids(el: Element, tags: string[]): Element[] {
  const kids = elementKids(el);
  if (kids.length !== tags.length) return notConvertible();
  kids.forEach((k, i) => { if (k.tagName !== tags[i]) notConvertible(); });
  return kids;
}

function runPiece(r: Element, guard = false): LatexPiece {
  onlyAttrs(r, []);
  let text: string | null = null;
  let preserve = false;
  let lit = false;
  for (const k of elementKids(r)) {
    if (k.tagName === 'm:rPr') {
      onlyAttrs(k, []);
      const pk = elementKids(k);
      if (pk.length !== 1 || pk[0].tagName !== 'm:lit') notConvertible();
      onlyAttrs(pk[0], []);
      lit = true;
    } else if (k.tagName === 'm:t') {
      if (text !== null) notConvertible();
      onlyAttrs(k, ['xml:space']);
      const sp = k.getAttribute('xml:space');
      if (sp !== null && sp !== '' && sp !== 'preserve') notConvertible();
      preserve = sp === 'preserve';
      for (let n = k.firstChild; n; n = n.nextSibling) if (n.nodeType !== 3) notConvertible();
      text = k.textContent ?? '';
    } else notConvertible();
  }
  if (text === null || text === '') return notConvertible();
  if (lit) {
    if (preserve) notConvertible();
    if (FUNCS.has(text)) return piece('\\' + text, false, false, true);
    if (/^[A-Za-z]+$/.test(text)) return piece(`\\operatorname{${text}}`);
    return notConvertible();
  }
  if (preserve) {
    if (/[{}]/.test(text)) notConvertible();
    return piece(`\\text{${text}}`);
  }
  // A bare run never holds whitespace or structural characters (the converter drops or splits them).
  if (/[\s\\|\]]/.test(text)) notConvertible();
  const chars = Array.from(text);
  if (chars.length === 1) {
    const c = chars[0];
    if ('%#$&_^{}'.includes(c)) return piece('\\' + c);
    if (NARY_NAME_BY_CHR[c]) return { ...piece('\\' + NARY_NAME_BY_CHR[c], false, false, true), endsNary: true };
    if (SYMBOL_NAME_BY_CHAR[c]) return piece('\\' + SYMBOL_NAME_BY_CHAR[c], false, false, true);
    return piece(guard ? `{${c}}` : c, true, true);
  }
  // Several characters stay one run, written as typed. Plain-text math (a formula with no backslash) is stored this
  // way, so `E=mc^2` exports as `E=mc^2`; only the characters that would break the LaTeX are escaped.
  const raw = chars.map((c) => ('&{}'.includes(c) ? '\\' + c : c)).join('');
  return piece(guard ? `{${raw}}` : raw, true, true);
}

/**
 * `safeEnd` is false in a matrix cell, where the parser looks for an operand after a trailing sum, product
 * or integral sign and fails at `&`, `\\\\` or `\\end`: that last piece is fenced in braces, which end the operand search.
 */
function latexSeq(kids: Element[], depth: number, safeEnd = true): string {
  if (depth > LATEX_MAX_DEPTH) notConvertible();
  let out = '';
  let prev: LatexPiece | null = null;
  kids.forEach((k, index) => {
    // A bare run right after a bare run would fuse into one run: the piece fences its own text.
    const p = latexPiece(k, depth + 1, prev !== null && prev.endsWord);
    let s = p.s;
    const last = index === kids.length - 1;
    // A sum or integral with both limits swallows the next operand as its integrand: fence it unless it is last.
    if ((p.fenceIfFollowed && !last) || (p.endsNary && last && !safeEnd)) s = `{${s}}`;
    else if (prev && prev.endsCmd && /^[A-Za-z]/.test(s)) s = ' ' + s;
    out += s;
    prev = { ...p, s };
  });
  return out;
}

function latexCell(el: Element, depth: number): string {
  const kids = elementKids(el);
  if (kids.length === 0) notConvertible();
  return latexSeq(kids, depth, false);
}

/** Content that must not be empty (the converter refuses empty groups). */
function latexSlot(el: Element, depth: number): string {
  const kids = elementKids(el);
  if (kids.length === 0) notConvertible();
  return latexSeq(kids, depth);
}

function delimiterText(el: Element): string {
  onlyAttrs(el, ['m:val']);
  const v = el.getAttribute('m:val');
  if (v === null) return notConvertible();
  if (v === '') return '.';
  if (v === '{') return '\\{';
  if (v === '}') return '\\}';
  if ('()[]|'.includes(v) && v.length === 1) return v;
  return notConvertible();
}

/**
 * A sub- or superscript operand. One run that is a whole operand (`2`, `n`, `10`, `\\alpha`) is written bare, which
 * keeps the converter's nesting depth equal to the original; anything else goes in braces.
 */
function latexLimit(el: Element, depth: number): LatexPiece {
  const kids = elementKids(el);
  if (kids.length === 0) notConvertible();
  if (kids.length === 1 && kids[0].tagName === 'm:r') {
    const p = runPiece(kids[0]);
    const text = kids[0].textContent ?? '';
    const wholeOperand = p.startsWord
      ? p.s === text && !text.includes('[') && operandPrefixLength(text) === text.length
      : !p.endsNary && !p.s.startsWith('\\text{') && !p.s.startsWith('\\operatorname{');
    if (wholeOperand) return p;
  }
  return piece(`{${latexSeq(kids, depth)}}`);
}

function latexPiece(el: Element, depth: number, guard = false): LatexPiece {
  if (depth > LATEX_MAX_DEPTH) notConvertible();
  switch (el.tagName) {
    case 'm:r':
      return runPiece(el, guard);
    case 'm:f': {
      onlyAttrs(el, []);
      const [num, den] = exactKids(el, ['m:num', 'm:den']);
      return piece(`\\frac{${latexSlot(num, depth)}}{${latexSlot(den, depth)}}`);
    }
    case 'm:rad': {
      onlyAttrs(el, []);
      const [pr, deg, e] = exactKids(el, ['m:radPr', 'm:deg', 'm:e']);
      const [hide] = exactKids(pr, ['m:degHide']);
      onlyAttrs(hide, ['m:val']);
      const hv = hide.getAttribute('m:val');
      const hasDegree = elementKids(deg).length > 0;
      if (hv === '1' && !hasDegree) return piece(`\\sqrt{${latexSlot(e, depth)}}`);
      if (hv === '0' && hasDegree) return piece(`\\sqrt[${latexSlot(deg, depth)}]{${latexSlot(e, depth)}}`);
      return notConvertible();
    }
    case 'm:acc': {
      onlyAttrs(el, []);
      const [pr, e] = exactKids(el, ['m:accPr', 'm:e']);
      const [chr] = exactKids(pr, ['m:chr']);
      onlyAttrs(chr, ['m:val']);
      const cmd = Object.keys(ACCENTS).find(k => ACCENTS[k] === chr.getAttribute('m:val'));
      if (!cmd) return notConvertible();
      return piece(`\\${cmd}{${latexSlot(e, depth)}}`);
    }
    case 'm:sSub':
    case 'm:sSup':
    case 'm:sSubSup': {
      onlyAttrs(el, []);
      const tags = el.tagName === 'm:sSub' ? ['m:e', 'm:sub'] : el.tagName === 'm:sSup' ? ['m:e', 'm:sup'] : ['m:e', 'm:sub', 'm:sup'];
      const kids = exactKids(el, tags);
      const baseKids = elementKids(kids[0]);
      if (baseKids.length === 0) notConvertible();
      // One run or one command-like node is written bare (x^{2}, \\frac{a}{b}^{2}); anything else is fenced
      // so the limits attach to all of it.
      const only = baseKids.length === 1 ? baseKids[0] : null;
      // A run that starts with a relation and has more text must stay fenced,
      // because a bare `=ab_i` now reads back as `=` followed by base `ab`.
      const relationLead = only !== null && only.tagName === 'm:r' && /^[=<>]+./s.test(only.textContent ?? '');
      const bare = only !== null && !relationLead && ['m:r', 'm:f', 'm:rad', 'm:d', 'm:m'].includes(only.tagName);
      const basePiece = bare ? latexPiece(only as Element, depth + 1, guard) : null;
      let s = basePiece ? basePiece.s : `{${latexSeq(baseKids, depth)}}`;
      let tail: LatexPiece | null = null;
      for (let i = 1; i < kids.length; i++) {
        tail = latexLimit(kids[i], depth);
        s += (kids[i].tagName === 'm:sub' ? '_' : '^') + tail.s;
      }
      const naryBase = only !== null && only.tagName === 'm:r' && NARY_NAME_BY_CHR[only.textContent ?? ''] !== undefined;
      const ends = tail as LatexPiece;
      if (naryBase) return { ...piece(s, false, ends.endsWord, ends.endsCmd), fenceIfFollowed: el.tagName === 'm:sSubSup', endsNary: true };
      return piece(s, basePiece !== null && basePiece.startsWord, ends.endsWord, ends.endsCmd);
    }
    case 'm:nary': {
      onlyAttrs(el, []);
      const [pr, sub, sup, e] = exactKids(el, ['m:naryPr', 'm:sub', 'm:sup', 'm:e']);
      onlyAttrs(pr, []);
      const [chr, lim] = exactKids(pr, ['m:chr', 'm:limLoc']);
      onlyAttrs(chr, ['m:val']);
      onlyAttrs(lim, ['m:val']);
      const name = NARY_NAME_BY_CHR[chr.getAttribute('m:val') ?? ''];
      if (name === undefined || lim.getAttribute('m:val') !== NARY[name].limLoc) return notConvertible();
      const head = `\\${name}_${latexLimit(sub, depth).s}^${latexLimit(sup, depth).s}`;
      const eKids = elementKids(e);
      // `\int_a^b \,` makes the converter emit an empty integrand (the spacing command is the operand).
      if (eKids.length === 0) return piece(`${head} \\,`);
      // The integrand is one operand: one node is written bare, several are fenced.
      if (eKids.length === 1) {
        const ip = latexPiece(eKids[0], depth + 1);
        // An integrand that is itself a trailing sum sign would take the next sibling as its operand: fence the pair.
        return {
          ...piece(`${head} ${ip.s}`, false, ip.endsWord, ip.endsCmd),
          endsNary: ip.endsNary,
          fenceIfFollowed: ip.endsNary === true || ip.fenceIfFollowed === true,
        };
      }
      return piece(`${head}{${latexSeq(eKids, depth)}}`);
    }
    case 'm:d': {
      onlyAttrs(el, []);
      const [pr, e] = exactKids(el, ['m:dPr', 'm:e']);
      onlyAttrs(pr, []);
      const [beg, end] = exactKids(pr, ['m:begChr', 'm:endChr']);
      return piece(`\\left${delimiterText(beg)}${latexSlot(e, depth)}\\right${delimiterText(end)}`);
    }
    case 'm:m': {
      onlyAttrs(el, []);
      const rows = elementKids(el);
      if (rows.length === 0) notConvertible();
      const lines = rows.map((mr) => {
        if (mr.tagName !== 'm:mr') notConvertible();
        onlyAttrs(mr, []);
        const cells = elementKids(mr);
        if (cells.length === 0) notConvertible();
        return cells.map((c) => {
          if (c.tagName !== 'm:e') notConvertible();
          onlyAttrs(c, []);
          return latexCell(c, depth);
        }).join(' & ');
      });
      return piece(`\\begin{matrix}${lines.join(' \\\\ ')}\\end{matrix}`);
    }
    default:
      return notConvertible();
  }
}

/**
 * Converts an `m:oMath` element back to LaTeX, or returns null when it holds anything
 * latexToOmml does not emit (callers then fall back to the flat `m:t` text).
 */
export function ommlToLatex(el: Element): string | null {
  try {
    if (!el || el.tagName !== 'm:oMath') return null;
    onlyAttrs(el, []);
    const kids = elementKids(el);
    if (kids.length === 0) return null;
    const out = latexSeq(kids, 0);
    return out === '' ? null : out;
  } catch {
    return null;
  }
}

// ── OMML → MathML (HTML export, display equations) ──────────────────────────
// Native MathML (MathML Core-compatible: no mfenced), well-formed by
// construction. Used for DISPLAY equations only; inline math keeps the
// legacy placeholder span.

const MATHML_NS = 'http://www.w3.org/1998/Math/MathML';

function escXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function mathKids(el: Element): string {
  let out = '';
  for (let i = 0; i < el.childNodes.length; i++) {
    const ch = el.childNodes[i] as Element;
    if (!ch || (ch as unknown as Node).nodeType !== 1) continue;
    out += mathNode(ch);
  }
  return out;
}

function firstChildEl(parent: Element, tag: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const ch = parent.childNodes[i] as Element;
    if (ch && (ch as unknown as Node).nodeType === 1 && (ch as Element).tagName === tag) {
      return ch as Element;
    }
  }
  return null;
}

function attrVal(el: Element, name: string): string {
  return el.getAttribute(name) || el.getAttribute(`m:${name}`) || '';
}

function mathRun(r: Element): string {
  let text = '';
  let lit = false;
  let textMode = false;
  for (let i = 0; i < r.childNodes.length; i++) {
    const ch = r.childNodes[i] as Element;
    if (!ch || (ch as unknown as Node).nodeType !== 1) continue;
    if (ch.tagName === 'm:t') {
      text += ch.textContent || '';
      if (ch.getAttribute('xml:space') === 'preserve') textMode = true;
    } else if (ch.tagName === 'm:rPr') {
      for (let k = 0; k < ch.childNodes.length; k++) {
        const g = ch.childNodes[k] as Element;
        if (g && (g as unknown as Node).nodeType === 1 && (g as Element).tagName === 'm:lit') lit = true;
      }
    }
  }
  if (text === '') return '';
  const esc = escXml(text);
  if (textMode) return `<mtext>${esc}</mtext>`;
  if (lit) return `<mi>${esc}</mi>`;
  if (/^[0-9.,]+$/.test(text)) return `<mn>${esc}</mn>`;
  // Split into letter runs (mi), number runs (mn), operator chars (mo).
  let out = '';
  let cur = '';
  let curKind = '';
  const push = (): void => {
    if (cur === '') return;
    const tag = curKind === 'num' ? 'mn' : curKind === 'id' ? 'mi' : 'mo';
    out += `<${tag}>${escXml(cur)}</${tag}>`;
    cur = '';
  };
  for (const ch of text) {
    const kind = /[0-9.]/.test(ch) ? 'num' : /[a-zA-Zα-ωΑ-Ωϵϑϖϱςϕℓ∂∇∞∈μ]/.test(ch) ? 'id' : 'op';
    if (kind !== curKind) push();
    curKind = kind;
    cur += ch;
  }
  push();
  return out;
}

function mathNode(el: Element): string {
  const tag = el.tagName || '';
  switch (tag) {
    case 'm:oMath':
    case 'm:oMathPara':
    case 'm:e':
    case 'm:num':
    case 'm:den':
    case 'm:deg':
    case 'm:sub':
    case 'm:sup':
    case 'm:lim':
      return `<mrow>${mathKids(el)}</mrow>`;
    case 'm:f': {
      const num = firstChildEl(el, 'm:num');
      const den = firstChildEl(el, 'm:den');
      return `<mfrac><mrow>${num ? mathKids(num) : ''}</mrow><mrow>${den ? mathKids(den) : ''}</mrow></mfrac>`;
    }
    case 'm:rad': {
      const deg = firstChildEl(el, 'm:deg');
      const e = firstChildEl(el, 'm:e');
      const eInner = e ? mathKids(e) : '';
      const degInner = deg ? mathKids(deg) : '';
      if (degInner === '') return `<msqrt><mrow>${eInner}</mrow></msqrt>`;
      return `<mroot><mrow>${eInner}</mrow><mrow>${degInner}</mrow></mroot>`;
    }
    case 'm:sSup': {
      const e = firstChildEl(el, 'm:e');
      const sup = firstChildEl(el, 'm:sup');
      return `<msup><mrow>${e ? mathKids(e) : ''}</mrow><mrow>${sup ? mathKids(sup) : ''}</mrow></msup>`;
    }
    case 'm:sSub': {
      const e = firstChildEl(el, 'm:e');
      const sub = firstChildEl(el, 'm:sub');
      return `<msub><mrow>${e ? mathKids(e) : ''}</mrow><mrow>${sub ? mathKids(sub) : ''}</mrow></msub>`;
    }
    case 'm:sSubSup': {
      const e = firstChildEl(el, 'm:e');
      const sub = firstChildEl(el, 'm:sub');
      const sup = firstChildEl(el, 'm:sup');
      return `<msubsup><mrow>${e ? mathKids(e) : ''}</mrow><mrow>${sub ? mathKids(sub) : ''}</mrow><mrow>${sup ? mathKids(sup) : ''}</mrow></msubsup>`;
    }
    case 'm:nary': {
      const pr = firstChildEl(el, 'm:naryPr');
      const chrEl = pr ? firstChildEl(pr, 'm:chr') : null;
      const limEl = pr ? firstChildEl(pr, 'm:limLoc') : null;
      const chr = chrEl ? (attrVal(chrEl, 'val') || '∑') : '∑';
      const loc = limEl ? attrVal(limEl, 'val') : 'undOvr';
      const sub = firstChildEl(el, 'm:sub');
      const sup = firstChildEl(el, 'm:sup');
      const e = firstChildEl(el, 'm:e');
      const subHtml = `<mrow>${sub ? mathKids(sub) : ''}</mrow>`;
      const supHtml = `<mrow>${sup ? mathKids(sup) : ''}</mrow>`;
      const eHtml = `<mrow>${e ? mathKids(e) : ''}</mrow>`;
      const op = `<mo>${escXml(chr)}</mo>`;
      const opTag = loc === 'undOvr'
        ? `<munderover>${op}${subHtml}${supHtml}</munderover>`
        : `<msubsup>${op}${subHtml}${supHtml}</msubsup>`;
      return `<mrow>${opTag}${eHtml}</mrow>`;
    }
    case 'm:d': {
      const pr = firstChildEl(el, 'm:dPr');
      const beg = pr ? firstChildEl(pr, 'm:begChr') : null;
      const end = pr ? firstChildEl(pr, 'm:endChr') : null;
      const open = beg ? attrVal(beg, 'val') : '';
      const close = end ? attrVal(end, 'val') : '';
      const e = firstChildEl(el, 'm:e');
      const inner = e ? mathKids(e) : '';
      const l = open !== '' ? `<mo>${escXml(open)}</mo>` : '';
      const r = close !== '' ? `<mo>${escXml(close)}</mo>` : '';
      return `<mrow>${l}${inner}${r}</mrow>`;
    }
    case 'm:m': {
      let rows = '';
      for (let i = 0; i < el.childNodes.length; i++) {
        const mr = el.childNodes[i] as Element;
        if (!mr || (mr as unknown as Node).nodeType !== 1 || (mr as Element).tagName !== 'm:mr') continue;
        let cells = '';
        for (let k = 0; k < mr.childNodes.length; k++) {
          const ce = mr.childNodes[k] as Element;
          if (!ce || (ce as unknown as Node).nodeType !== 1 || (ce as Element).tagName !== 'm:e') continue;
          cells += `<mtd>${mathKids(ce)}</mtd>`;
        }
        rows += `<mtr>${cells}</mtr>`;
      }
      return `<mtable>${rows}</mtable>`;
    }
    case 'm:r':
      return mathRun(el);
    case 'm:t':
      return `<mtext>${escXml(el.textContent || '')}</mtext>`;
    case 'm:rPr':
    case 'm:naryPr':
    case 'm:dPr':
    case 'm:radPr':
    case 'm:mPr':
    case 'm:chr':
    case 'm:limLoc':
    case 'm:begChr':
    case 'm:endChr':
    case 'm:degHide':
    case 'm:ctrlPr':
      return '';
    default: {
      // Unknown wrappers: recurse so text survives; text-only leaves → mtext.
      const inner = mathKids(el);
      if (inner !== '') return `<mrow>${inner}</mrow>`;
      const t = el.textContent || '';
      return t !== '' ? `<mtext>${escXml(t)}</mtext>` : '';
    }
  }
}

/**
 * Renders an equation subtree (w:p equation paragraph, m:oMathPara, or
 * m:oMath) as a native MathML `<math display="block">` string. Returns ''
 * when no math content is found (caller falls back to the placeholder).
 */
export function ommlToMathmlDisplay(xml: Element | null | undefined): string {
  if (!xml || typeof (xml as unknown as { getElementsByTagName?: unknown }).getElementsByTagName !== 'function') {
    return '';
  }
  const roots: Element[] = [];
  const tag = (xml as Element).tagName || '';
  if (tag === 'm:oMath' || tag === 'm:oMathPara') {
    roots.push(xml as Element);
  } else {
    const paras = Array.from((xml as Element).getElementsByTagName('m:oMathPara'));
    for (const p of paras) roots.push(p as Element);
    if (roots.length === 0) {
      const maths = Array.from((xml as Element).getElementsByTagName('m:oMath'));
      for (const m of maths) roots.push(m as Element);
    }
  }
  if (roots.length === 0) return '';
  let out = '';
  for (const r of roots) {
    out += `<math xmlns="${MATHML_NS}" display="block">${mathKids(r)}</math>`;
  }
  return out;
}
