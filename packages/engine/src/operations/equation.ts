import { createWordElement, createMathElement } from '../ooxml/ns';
import { toError } from '../errors';
import { latexToOmml } from '../ooxml/omml';
import { DocumentAST, Node, Section, calculateNodeChecksum } from '../ast';
import { BaseOperation, InsertChainHost } from './base';
import { XMLSerializer, DOMParser } from '@xmldom/xmldom';
import * as crypto from 'crypto';
import { sanitizeXml } from '../validator';
import { DocxParser } from '../parser';
import { insertIntoSection } from '../parser/sections';
import { resolveReproducibleMode } from '../zip-hardening';
import { assertNoXxe } from '../ooxml/xxe-guard';

export interface EquationOperation extends BaseOperation {
  action: 'equation';
  operation: 'insert' | 'replace';
  anchor?: string;
  position?: 'before' | 'after';
  target?: string;
  /**
   * E4 doc line: `latex` is a LaTeX SUBSET (converted to structural OMML) or
   * plain math text (no backslashes — stored verbatim, unchanged).
   * Supported: fractions, roots, sub/superscripts, n-ary operators,
   * delimiters, matrix environments, symbol table, named functions.
   * Unsupported constructs (alignment envs, exotic macros, over-depth
   * nesting past 6) fail loudly via E_VALIDATION, never silently.
   */
  latex: string;
  /**
   * `auto` (default): LaTeX containing a backslash, `^` or `_` becomes a real
   * equation, anything else stays literal text. `latex`: always convert.
   * `plain`: keep the text exactly as written.
   */
  mode?: 'auto' | 'latex' | 'plain';
  id?: string;
  params?: Record<string, unknown>;
}

export interface EquationResult {
  id: string;
}

// Module-level counter so generated ids stay unique even within the same millisecond.
let equationIdSeq = 0;

function isEquationReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function equationHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

export function assertSafeLatex(latex: string): void {
  assertNoXxe(latex, 'equation latex');
  if (/<script/i.test(latex)) {
    throw new Error('Security Error: Executable script tags are prohibited in OOXML.');
  }
}

/**
 * OX-02 DECISION: IMPLEMENT (minimal Office Math), not remove.
 *
 * Why implement: the `equation` route + `Node.type='equation'` + parser
 * `hasEquation` flags + TOON `[EQ]` surface + validator equation-count layer
 * + README/test-llm caller contracts all assume equations exist. Removing
 * would require touching `backend/ast.ts`, `backend/parser.ts`,
 * `backend/toon.ts`, `backend/validator.ts`, docs/prompts — all outside D2
 * file ownership. A minimal OMML wrapper (`w:p > m:oMathPara > m:oMath >
 * m:r > m:t`) is genuine Word-compatible math (renders in Word, survives the
 * validator's per-node equation-count check on later edits, round-trips
 * through rezip) and un-poisons sequential batches that previously always
 * threw `equation.* is not supported yet`.
 *
 * E4 (convert-then-error, graduated): `latex` is a LaTeX SUBSET (converted to
 * structural OMML by latexToOmml in ooxml/omml.ts) or plain math text
 * (stored verbatim, unchanged). Backslash-bearing input runs through the
 * E4 converter (symbol table → Unicode, \frac→m:f, \sqrt→m:rad,
 * ^/_→m:sSup/m:sSub/m:sSubSup, \sum/\int→m:nary, \left..\right/|..|→m:d,
 * matrix-family \begin{...}→m:m, \text/\mbox→literal runs, recursive to
 * depth 6); anything the converter cannot handle (alignment environments,
 * exotic macros, over-depth nesting) throws E_VALIDATION naming the
 * construct + position + remedy, and converter output is re-validated
 * through the real validator path (sanitizeXml XXE/script gates + strict
 * re-parse + m:oMath presence) so half-converted XML is never emitted.
 * Plain-text math with NO backslashes keeps the original verbatim behavior
 * byte-identical (existing suites pin it). assertSafeLatex/XXE gates stay
 * on the input AND cover the output.
 */
function buildEquationParagraph(doc: Document, id: string, latex: string, mode?: EquationOperation['mode']): Element {
  const p = createWordElement(doc, 'w:p');
  p.setAttribute('data-astra-id', id);
  p.appendChild(buildValidatedOMathPara(doc, latex, mode));
  // Real validator path on the produced paragraph: never emit half-converted XML.
  assertValidConvertedParagraph(p);
  return p;
}

// ── E4 converter: LaTeX subset → structural OMML (packages/engine/src/ooxml/omml.ts) ──
// Tokenizer + recursive DOM builder (OMML_MAX_DEPTH = 6, loud beyond).
// Strict superset of the retired E1 mini-converter: every E1-mapped macro
// maps identically, plus matrix environments (m:m), a larger symbol table,
// \det, \quad/\qquad spacing, and \begin{array}{lcr}.

/** Builds the OMML body shared by insert + replace (validated before any DOM/AST mutation). */
function buildValidatedOMathPara(doc: Document, latex: string, mode?: EquationOperation['mode']): Element {
  const oMathPara = createMathElement(doc, 'm:oMathPara');
  const hasBackslash = latex.includes('\\');
  const treatAsLatex = mode === 'latex' || (mode !== 'plain' && /[\\^_]/.test(latex));
  if (treatAsLatex) {
    try {
      oMathPara.appendChild(latexToOmml(doc, latex));
      return oMathPara;
    } catch (err) {
      // auto mode only: text without a backslash that the converter rejects (for
      // example `x^`) keeps verbatim storage instead of becoming a new error.
      if (mode === 'latex' || hasBackslash) throw err;
    }
  }
  // Plain-text math: stored verbatim.
  const oMath = createMathElement(doc, 'm:oMath');
  const r = createMathElement(doc, 'm:r');
  const t = createMathElement(doc, 'm:t');
  t.appendChild(doc.createTextNode(latex));
  r.appendChild(t);
  oMath.appendChild(r);
  oMathPara.appendChild(oMath);
  return oMathPara;
}

/**
 * Real validator path for converter output: XXE/script gates (extended from
 * the input gate) + strict re-parse + m:oMath presence. Anything not
 * validating throws E_VALIDATION — half-converted XML is never emitted.
 */
export function assertValidConvertedParagraph(p: Element): void {
  const xml = new XMLSerializer().serializeToString(p);
  try {
    sanitizeXml(xml);
    const strict = new DOMParser({
      errorHandler: {
        warning: (w) => { throw new Error(String(w)); },
        error: (e) => { throw new Error(String(e)); },
        fatalError: (f) => { throw new Error(String(f)); },
      },
    });
    const dom = strict.parseFromString(xml, 'text/xml');
    const root = dom.documentElement;
    if (!root || root.tagName !== 'w:p') throw new Error('converted paragraph root is not w:p');
    if (dom.getElementsByTagName('m:oMath').length < 1) throw new Error('converted paragraph lost its math content');
  } catch (err) {
    if (err instanceof Error && err.message.includes('E_VALIDATION')) throw err;
    const msg = err instanceof Error ? err.message : String(err);
    const wrapped = new Error(
      `E_VALIDATION: equation converter produced invalid OMML (${msg}). Remedy: simplify the expression, split it into smaller equations, or insert pre-built OMML instead of LaTeX.`
    );
    (wrapped as unknown as Record<string, unknown>).code = 'E_VALIDATION';
    (wrapped as unknown as Record<string, unknown>).retryable = false;
    throw wrapped;
  }
}

function findNode(ast: DocumentAST, id: string): { section: Section; index: number; node: Node } | null {
  for (const section of ast.sections) {
    const index = section.nodes.findIndex(n => n.id === id);
    if (index !== -1) return { section, index, node: section.nodes[index] };
  }
  return null;
}

export async function executeEquation(
  op: EquationOperation,
  ast: DocumentAST,
  parser?: DocxParser
): Promise<EquationResult> {
  // OP-16: accept args nested under params:{} or flat (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedOperation = String(pick('operation', op.operation) ?? op.operation ?? '');
  const resolvedAnchor = pick('anchor', op.anchor) as string | undefined;
  const resolvedTarget = pick('target', op.target) as string | undefined;
  const resolvedLatex = pick('latex', op.latex) as string | undefined;
  const resolvedPosition = (pick('position', op.position) ?? op.position ?? 'after') as 'before' | 'after';
  const customId = pick('id', op.id) as string | undefined;
  const resolvedMode = pick('mode', op.mode) as EquationOperation['mode'];

  if (resolvedOperation !== 'insert' && resolvedOperation !== 'replace') {
    throw new Error(
      `equation.${resolvedOperation || '(missing)'} is not supported yet. Supported: insert, replace. ` +
      `Re-read the current structure with inspect_document and retry with operation 'insert' or 'replace'.`
    );
  }
  if (typeof resolvedLatex !== 'string' || resolvedLatex.trim() === '') {
    throw new Error(
      'Equation operation requires non-empty `latex`. ' +
      'Re-read the current structure with inspect_document and retry with the LaTeX source in `latex`.'
    );
  }
  const latex = resolvedLatex;
  assertSafeLatex(latex);

  if (resolvedOperation === 'insert') {
    if (!resolvedAnchor) throw new Error("Insert equation requires an anchor. Re-read the current structure with inspect_document and retry with a valid node id.");
    const found = findNode(ast, resolvedAnchor);
    if (!found) {
      throw new Error(
        `Anchor node '${resolvedAnchor}' not found for equation insertion. ` +
        `Re-read the current structure with inspect_document and retry with a valid node id.`
      );
    }
    const doc = found.node.xml.ownerDocument;
    if (!doc) throw new Error(`Anchor node '${resolvedAnchor}' has no owner document; cannot insert equation`);
    if (resolvedPosition !== 'before' && resolvedPosition !== 'after') {
      throw new Error(
        `Unknown equation insert position '${resolvedPosition}'. Supported: before, after. ` +
        `Re-read the current structure with inspect_document and retry.`
      );
    }

    const newId = customId || (isEquationReproducible()
      ? (() => {
          const base = `_eq_${equationHash8(`${resolvedAnchor ?? ''}:${resolvedPosition}:${latex}`)}`;
          const live = new Set<string>();
          try {
            for (const s of ast.sections ?? []) {
              for (const n of s.nodes ?? []) {
                if (typeof n.id === 'string' && n.id) live.add(n.id);
              }
            }
          } catch { /* best-effort */ }
          let candidate = base;
          let n = 2;
          while (live.has(candidate)) candidate = `${base}_${n++}`;
          return candidate;
        })()
      : `_eq_new_${Date.now()}_${++equationIdSeq}`);
    const p = buildEquationParagraph(doc, newId, latex, resolvedMode);

    // Same-anchor 'after' inserts preserve FIFO order via a per-anchor
    // chain (mirrors insert.ts): re-validated against live DOM/AST at use,
    // dead links dropped so the next insert lands after the latest live
    // link, never after a detached ghost.
    let astSection: Section = found.section;
    let astInsertIndex = resolvedPosition === 'before' ? found.index : found.index + 1;
    let chainHost: InsertChainHost | null = null;
    let liveChain: Node[] = [];
    if (resolvedPosition === 'before') {
      found.node.xml.parentNode?.insertBefore(p, found.node.xml);
    } else {
      chainHost = found.node as unknown as InsertChainHost;
      const rawChain: Node[] = Array.isArray(chainHost._insertAfterChain)
        ? (chainHost._insertAfterChain as Node[])
        : [];
      const legacyTail = chainHost._lastInsertAfter as Node | undefined;
      if (rawChain.length === 0 && legacyTail) {
        rawChain.push(legacyTail);
      }
      liveChain = rawChain.filter((n) => {
        if (!n || !n.xml || !(n.xml as Element).parentNode) return false;
        return ast.sections.some((s: Section) => s.nodes.some((x: Node) => x.id === n.id));
      });
      chainHost._insertAfterChain = liveChain;
      if (legacyTail && !liveChain.some((n) => n.id === legacyTail.id)) {
        delete chainHost._lastInsertAfter;
      }
      const effectiveAnchor: Node = liveChain.length > 0 ? liveChain[liveChain.length - 1] : found.node;
      const effXml = effectiveAnchor.xml as Element;
      effXml.parentNode?.insertBefore(p, effXml.nextSibling);
      let effSection: Section = found.section;
      let effIdx = (found.section.nodes as Node[]).findIndex((n: Node) => n.id === effectiveAnchor.id);
      if (effIdx === -1) {
        for (const s of ast.sections) {
          const i = (s.nodes as Node[]).findIndex((n: Node) => n.id === effectiveAnchor.id);
          if (i !== -1) {
            effSection = s;
            effIdx = i;
            break;
          }
        }
      }
      astSection = effSection;
      astInsertIndex = effIdx !== -1 ? effIdx + 1 : found.index + 1;
    }

    const xmlStr = new XMLSerializer().serializeToString(p);
    const newNode: Node = {
      id: newId,
      type: 'equation',
      xml: p,
      metadata: {
        text: latex,
        hasCitation: false,
        hasComment: false,
        hasBookmark: false,
        hasTrackedChanges: false,
        hasEquation: true,
        hasField: false,
        hasImage: false,
        hasHyperlink: false,
      },
      version: 1,
      checksum: await calculateNodeChecksum(xmlStr),
    };
    const insertIndex = astInsertIndex;
    insertIntoSection(ast.sections, astSection, insertIndex, [newNode]);
    if (chainHost) {
      liveChain.push(newNode);
      chainHost._insertAfterChain = liveChain;
      chainHost._lastInsertAfter = newNode;
    }
    try {
      parser?.nodeMap?.set?.(newId, p);
    } catch { /* best-effort — AST is the source of truth here */ }
    return { id: newId };
  }

  // replace
  if (!resolvedTarget) throw new Error("Replace equation requires a target. Re-read the current structure with inspect_document and retry with a valid node id.");
  const found = findNode(ast, resolvedTarget);
  if (!found) {
    throw new Error(
      `Node '${resolvedTarget}' not found for equation replacement. ` +
      `Re-read the current structure with inspect_document and retry with a valid node id.`
    );
  }
  const doc = found.node.xml.ownerDocument;
  if (!doc) throw new Error(`Node '${resolvedTarget}' has no owner document; cannot replace equation`);

  // Preserve paragraph properties (pPr) so a styled paragraph keeps its
  // style; drop other children and append the OMML run.
  // E1: build + validate the replacement FIRST so a loud converter failure
  // never leaves a half-converted target behind.
  const probe = createWordElement(doc, 'w:p');
  const replacement = buildValidatedOMathPara(doc, latex, resolvedMode);
  probe.appendChild(replacement);
  assertValidConvertedParagraph(probe);
  const targetEl = found.node.xml as Element;
  if (targetEl.tagName !== 'w:p') {
    throw toError(
      'E_VALIDATION',
      `Cannot replace equation on node '${resolvedTarget}' of type '${found.node.type}' (tag ${targetEl.tagName}). Equation replace targets paragraph nodes.`,
      'Re-read the current structure with inspect_document and retry with a paragraph or equation node id.'
    );
  }
  // Replace keeps only the paragraph properties, so it would delete a picture, a chart, or a comment or note anchor.
  const holdsObject = ['w:drawing', 'w:pict', 'w:object', 'w:commentReference', 'w:footnoteReference', 'w:endnoteReference']
    .some((tag) => targetEl.getElementsByTagName(tag).length > 0);
  if (found.node.type === 'image' || found.node.type === 'chart' || found.node.type === 'chart3d' || holdsObject) {
    throw toError(
      'E_VALIDATION',
      `Cannot replace equation on node '${resolvedTarget}': it holds a picture, chart, comment or note anchor that the replacement would delete.`,
      "Insert the equation next to it with action 'equation' operation 'insert', or delete the object first."
    );
  }
  const pPr = targetEl.getElementsByTagName('w:pPr')[0];
  while (targetEl.firstChild) targetEl.removeChild(targetEl.firstChild);
  if (pPr) targetEl.appendChild(pPr);
  // appendChild MOVES the already-validated node out of the probe.
  targetEl.appendChild(replacement);

  found.node.metadata.text = latex;
  found.node.metadata.hasEquation = true;
  found.node.type = 'equation';
  found.node.version += 1;
  return { id: found.node.id };
}
