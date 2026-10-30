import { createWordElement } from '../ooxml/ns';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST } from '../ast';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { createHash } from 'crypto';

export const NUMBERING_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NUMBERING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
export const NUMBERING_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';

/** A list definition carries every level Word lets a list nest to (w:ilvl 0-8). */
export const LIST_LEVEL_COUNT = 9;

/** Word's own pattern: bullets cycle •, ◦, ▪; numbers cycle 1., a., i. and repeat. */
const BULLET_GLYPHS = ['•', '◦', '▪'];
const NUMBER_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman'];
/** Each level sits one 0.5 inch step further in, with a 0.25 inch hanging marker. */
const LEVEL_INDENT_STEP = 720;
const LEVEL_HANGING = 360;

type ListKind = 'bullet' | 'numbered';

interface ListLevelSpec {
  numFmt: string;
  lvlText: string;
  left: number;
  hanging: number;
}

/** What level `ilvl` (0-based, as in the file) of a new `kind` list holds. */
function listLevelSpec(kind: ListKind, ilvl: number): ListLevelSpec {
  const left = LEVEL_INDENT_STEP * (ilvl + 1);
  return kind === 'bullet'
    ? { numFmt: 'bullet', lvlText: BULLET_GLYPHS[ilvl % BULLET_GLYPHS.length], left, hanging: LEVEL_HANGING }
    : { numFmt: NUMBER_FORMATS[ilvl % NUMBER_FORMATS.length], lvlText: `%${ilvl + 1}.`, left, hanging: LEVEL_HANGING };
}

/** `<w:lvl>` markup for the fallback writer; the DOM builder below writes the same level. */
function listLevelXml(kind: ListKind, ilvl: number): string {
  const spec = listLevelSpec(kind, ilvl);
  const glyph = [...spec.lvlText].map((c) => (c.charCodeAt(0) > 126 ? `&#${c.charCodeAt(0)};` : c)).join('');
  return (
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${spec.numFmt}"/>` +
    `<w:lvlText w:val="${glyph}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind w:left="${spec.left}" w:hanging="${spec.hanging}"/></w:pPr>` +
    (kind === 'bullet' ? `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr>` : ``) +
    `</w:lvl>`
  );
}

/**
 * A definition's identity (`w:nsid`): eight hex digits that follow the content,
 * so the same request gives the same bytes, and that no other definition in
 * the document already uses.
 */
function listNsid(kind: ListKind, absId: string, taken: Set<string>): string {
  for (let attempt = 0; ; attempt++) {
    const nsid = createHash('sha1').update(`astra-list:${kind}:${absId}:${attempt}`).digest('hex').slice(0, 8).toUpperCase();
    if (!taken.has(nsid)) return nsid;
  }
}

/** Child elements of `el` with this tag; unlike getElementsByTagName it never reaches into nested elements. */
function directChildren(el: Element, tag: string): Element[] {
  const out: Element[] = [];
  for (let node = el.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && (node as Element).tagName === tag) out.push(node as Element);
  }
  return out;
}

export class NumberingManipulator {
  private ast: DocumentAST;

  constructor(ast: DocumentAST) {
    this.ast = ast;
  }

  /**
   * BX-03: parses word/numbering.xml into abstract + concrete maps keyed by
   * w:abstractNumId / w:numId. Tolerates missing/empty input (leaves maps as-is).
   * Elements keep their ownerDocument so later stubs can seed new nodes from it.
   */
  public parseNumbering(xmlStr: string): void {
    if (!xmlStr || typeof xmlStr !== 'string' || !xmlStr.includes('<')) return;
    assertNoXxe(xmlStr, 'word/numbering.xml');
    let doc: Document;
    try {
      doc = new DOMParser({
        errorHandler: {
          warning: () => {},
          error: (e: string) => { throw new Error(e); },
          fatalError: (f: string) => { throw new Error(f); },
        },
      }).parseFromString(xmlStr, 'text/xml');
    } catch {
      return;
    }
    if (!doc || !doc.documentElement) return;
    if (!this.ast.numbering) this.ast.numbering = { abstractNum: {}, concreteNum: {} };
    if (!this.ast.numbering.abstractNum) this.ast.numbering.abstractNum = {};
    if (!this.ast.numbering.concreteNum) this.ast.numbering.concreteNum = {};

    const abstracts = doc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < abstracts.length; i++) {
      const el = abstracts[i] as unknown as Element;
      const id = el.getAttribute('w:abstractNumId');
      if (id != null && id !== '') {
        (this.ast.numbering.abstractNum as Record<string, Element>)[id] = el as unknown as Element;
      }
    }
    const nums = doc.getElementsByTagName('w:num');
    for (let i = 0; i < nums.length; i++) {
      const el = nums[i] as unknown as Element;
      const id = el.getAttribute('w:numId');
      if (id != null && id !== '') {
        (this.ast.numbering.concreteNum as Record<string, Element>)[id] = el as unknown as Element;
      }
    }
  }

  // Ensures we reuse existing list definitions when possible to avoid XML bloat
  public getOrCreateList(type: 'bullet' | 'numbered'): string {
    // Returns a concrete numId
    // BX-10 (manipulator scope): unknown styles must fail loudly instead of
    // silently coercing to the numbered branch below.
    if (type !== 'bullet' && type !== 'numbered') {
      throw new Error(
        `Unknown list type '${String(type)}'. Expected 'bullet' or 'numbered'. ` +
        `Re-read the document with inspect_document and retry with a valid list type.`
      );
    }
    // If none exists, creates abstractNum and num, returning new numId
    const concrete = (this.ast.numbering?.concreteNum ?? {}) as Record<string, Element>;
    const abstract = (this.ast.numbering?.abstractNum ?? {}) as Record<string, Element>;

    // 1. Reuse an existing definition whose top level is the requested style.
    // Only level 1 decides: a Word list that is numbered at level 1 and bulleted
    // at level 2 is a numbered list, and must not be handed out for a bullet list.
    for (const [numId, numEl] of Object.entries(concrete)) {
      try {
        const absRef = (numEl as Element).getElementsByTagName?.('w:abstractNumId')?.[0] as Element | undefined;
        const absId = absRef?.getAttribute?.('w:val');
        if (absId == null) continue;
        const absEl = abstract[absId];
        if (!absEl) continue;
        if (this.topLevelKind(absEl) === type) return numId;
      } catch {
        continue;
      }
    }

    // 2. No reusable definition — allocate fresh ids (bullet and numbered must
    // not share a numId; the old stub returned '1' for both styles).
    const usedNumIds = new Set(
      Object.keys(concrete).map((k) => parseInt(k, 10)).filter((n) => !Number.isNaN(n))
    );
    const usedAbsIds = new Set(
      Object.keys(abstract).map((k) => parseInt(k, 10)).filter((n) => !Number.isNaN(n))
    );
    // Conventional defaults: bullet -> numId 1 / abstractNumId 0,
    // numbered -> numId 2 / abstractNumId 1. Bump past max when taken.
    const preferredNum = type === 'bullet' ? 1 : 2;
    let numId = preferredNum;
    if (usedNumIds.has(numId)) {
      numId = Math.max(0, ...[...usedNumIds, 0]) + 1;
    }
    let absId = preferredNum - 1;
    if (usedAbsIds.has(absId)) {
      absId = Math.max(0, ...[...usedAbsIds, 0]) + 1;
    }

    // 3. Record a schema-complete definition with every level, so a list item
    // placed at any level (1-9) finds its marker and indent, and future calls
    // reuse it without XML bloat. The bullet glyphs are plain Unicode in Arial:
    // Word's own Symbol/Wingdings code points are private-use and show tofu
    // where those fonts are missing.
    try {
      const seedDoc = this.findSeedDocument();
      if (seedDoc) {
        const absEl = createWordElement(seedDoc, 'w:abstractNum');
        absEl.setAttribute('w:abstractNumId', String(absId));
        const nsid = createWordElement(seedDoc, 'w:nsid');
        nsid.setAttribute('w:val', listNsid(type, String(absId), this.takenNsids()));
        absEl.appendChild(nsid);
        const multiLevelType = createWordElement(seedDoc, 'w:multiLevelType');
        multiLevelType.setAttribute('w:val', 'hybridMultilevel');
        absEl.appendChild(multiLevelType);
        for (let ilvl = 0; ilvl < LIST_LEVEL_COUNT; ilvl++) {
          absEl.appendChild(this.createLevel(seedDoc, type, ilvl));
        }
        const numEl = createWordElement(seedDoc, 'w:num');
        numEl.setAttribute('w:numId', String(numId));
        const absRefEl = createWordElement(seedDoc, 'w:abstractNumId');
        absRefEl.setAttribute('w:val', String(absId));
        numEl.appendChild(absRefEl);
        if (!this.ast.numbering.abstractNum) this.ast.numbering.abstractNum = {};
        if (!this.ast.numbering.concreteNum) this.ast.numbering.concreteNum = {};
        (this.ast.numbering.abstractNum as Record<string, Element>)[String(absId)] =
          absEl as unknown as Element;
        (this.ast.numbering.concreteNum as Record<string, Element>)[String(numId)] =
          numEl as unknown as Element;
      }
    } catch {
      // Non-fatal: still return the allocated id even if bookkeeping fails.
    }
    return String(numId);
  }

  /**
   * What level 1 of a definition is, or null when the definition is not a plain
   * list: it has no level 1, its level 1 shows no marker, or it belongs to a
   * paragraph style (heading numbering) or to another list's style, so handing
   * it to list items would tangle them with that style's numbers.
   */
  private topLevelKind(absEl: Element): ListKind | null {
    if (this.isStyleBound(absEl)) return null;
    const top = directChildren(absEl, 'w:lvl').find((lvl) => lvl.getAttribute('w:ilvl') === '0');
    if (!top) return null;
    const fmt = directChildren(top, 'w:numFmt')[0]?.getAttribute('w:val');
    if (fmt === 'bullet') return 'bullet';
    return fmt && fmt !== 'none' ? 'numbered' : null;
  }

  /** True for numbering that belongs to a paragraph style (headings) or to another list's style. */
  private isStyleBound(absEl: Element): boolean {
    return (
      directChildren(absEl, 'w:numStyleLink').length > 0 ||
      directChildren(absEl, 'w:styleLink').length > 0 ||
      directChildren(absEl, 'w:lvl').some((lvl) => directChildren(lvl, 'w:pStyle').length > 0)
    );
  }

  /** The `w:lvl` for level `ilvl` (0-based) of a new `kind` list, built in `doc`. */
  private createLevel(doc: Document, kind: ListKind, ilvl: number): Element {
    const spec = listLevelSpec(kind, ilvl);
    const lvl = createWordElement(doc, 'w:lvl');
    lvl.setAttribute('w:ilvl', String(ilvl));
    const withVal = (tag: string, val: string): Element => {
      const el = createWordElement(doc, tag);
      el.setAttribute('w:val', val);
      return el;
    };
    lvl.appendChild(withVal('w:start', '1'));
    lvl.appendChild(withVal('w:numFmt', spec.numFmt));
    lvl.appendChild(withVal('w:lvlText', spec.lvlText));
    lvl.appendChild(withVal('w:lvlJc', 'left'));
    const pPr = createWordElement(doc, 'w:pPr');
    const ind = createWordElement(doc, 'w:ind');
    ind.setAttribute('w:left', String(spec.left));
    ind.setAttribute('w:hanging', String(spec.hanging));
    pPr.appendChild(ind);
    lvl.appendChild(pPr);
    if (kind === 'bullet') {
      const rPr = createWordElement(doc, 'w:rPr');
      const fonts = createWordElement(doc, 'w:rFonts');
      fonts.setAttribute('w:ascii', 'Arial');
      fonts.setAttribute('w:hAnsi', 'Arial');
      fonts.setAttribute('w:cs', 'Arial');
      rPr.appendChild(fonts);
      lvl.appendChild(rPr);
    }
    return lvl;
  }

  /** Every `w:nsid` in use, upper-cased, so a new definition never repeats one. */
  private takenNsids(): Set<string> {
    const taken = new Set<string>();
    for (const absEl of Object.values((this.ast.numbering?.abstractNum ?? {}) as Record<string, Element>)) {
      const val = directChildren(absEl, 'w:nsid')[0]?.getAttribute('w:val');
      if (val) taken.add(val.toUpperCase());
    }
    return taken;
  }

  /**
   * Gives the definition behind `numId` every level 1-9 it lacks, in place, so
   * a list item placed at any level finds its marker and indent. Lists written
   * by earlier versions have one level, and short lists from Word can end at 3;
   * they upgrade the first time an item uses them. Levels already there, and
   * levels a `w:num` supplies through `w:lvlOverride`, are never touched. An
   * unknown `numId` is left alone, and so is numbering that belongs to a
   * paragraph style (heading numbering): its levels are the style's to define.
   */
  public ensureLevels(numId: string): void {
    const numEl = (this.ast.numbering?.concreteNum as Record<string, Element> | undefined)?.[numId];
    const absId = numEl?.getElementsByTagName?.('w:abstractNumId')?.[0]?.getAttribute?.('w:val');
    const absEl = absId != null ? (this.ast.numbering?.abstractNum as Record<string, Element> | undefined)?.[absId] : undefined;
    if (!numEl || !absEl || this.isStyleBound(absEl)) return;

    const present = new Map<number, Element>();
    for (const lvl of directChildren(absEl, 'w:lvl')) {
      const ilvl = parseInt(lvl.getAttribute('w:ilvl') ?? '', 10);
      if (Number.isInteger(ilvl)) present.set(ilvl, lvl);
    }
    const overridden = new Set<number>();
    for (const override of directChildren(numEl, 'w:lvlOverride')) {
      if (directChildren(override, 'w:lvl').length === 0) continue;
      const ilvl = parseInt(override.getAttribute('w:ilvl') ?? '', 10);
      if (Number.isInteger(ilvl)) overridden.add(ilvl);
    }
    const missing: number[] = [];
    for (let ilvl = 0; ilvl < LIST_LEVEL_COUNT; ilvl++) {
      if (!present.has(ilvl) && !overridden.has(ilvl)) missing.push(ilvl);
    }
    if (missing.length === 0) return;

    const doc = absEl.ownerDocument as unknown as Document | null;
    if (!doc) return;
    const lowest = [...present.entries()].sort(([a], [b]) => a - b)[0];
    const lowestFmt = lowest ? directChildren(lowest[1], 'w:numFmt')[0]?.getAttribute('w:val') : undefined;
    const kind: ListKind = lowestFmt && lowestFmt !== 'bullet' ? 'numbered' : 'bullet';
    for (const ilvl of missing) {
      const lvl = this.createLevel(doc, kind, ilvl);
      // Keep w:lvl in level order: before the first existing level above this one.
      const next = [...present.entries()].filter(([k]) => k > ilvl).sort(([a], [b]) => a - b)[0];
      if (next) absEl.insertBefore(lvl, next[1]);
      else absEl.appendChild(lvl);
      present.set(ilvl, lvl);
    }

    // A definition that nests is multilevel: a missing or single-level type would contradict its levels.
    const typeEl = directChildren(absEl, 'w:multiLevelType')[0];
    if (typeEl) {
      if (typeEl.getAttribute('w:val') === 'singleLevel') typeEl.setAttribute('w:val', 'hybridMultilevel');
    } else {
      const created = createWordElement(doc, 'w:multiLevelType');
      created.setAttribute('w:val', 'hybridMultilevel');
      const nsid = directChildren(absEl, 'w:nsid')[0];
      absEl.insertBefore(created, nsid ? nsid.nextSibling : absEl.firstChild);
    }
  }

  private findSeedDocument(): Document | null {
    for (const existing of Object.values(
      (this.ast.numbering?.concreteNum ?? {}) as Record<string, Element>
    )) {
      const doc = (existing as Element)?.ownerDocument;
      if (doc) return doc as unknown as Document;
    }
    for (const existing of Object.values(
      (this.ast.numbering?.abstractNum ?? {}) as Record<string, Element>
    )) {
      const doc = (existing as Element)?.ownerDocument;
      if (doc) return doc as unknown as Document;
    }
    for (const section of this.ast.sections ?? []) {
      for (const node of section.nodes ?? []) {
        const doc = (node.xml as Element)?.ownerDocument;
        if (doc) return doc as unknown as Document;
      }
    }
    return null;
  }

  // Handles complex task of continuing an interrupted list
  public continueList(previousNodeId: string): string {
    // BX-03: real continuation — looks up the numId of the previous list item
    // and returns it so the new item joins the same numbering sequence.
    // Throws loudly when the anchor is unknown or carries no list numbering
    // (the old stub unconditionally returned '1', forking every continued
    // list into the wrong sequence after save/reload).
    if (!previousNodeId || typeof previousNodeId !== 'string') {
      throw new Error(
        `continueList requires a previous list node id; received '${String(previousNodeId)}'. ` +
        `Re-read the document with inspect_document and retry with a valid node id.`
      );
    }
    let prev: { xml: Element } | undefined;
    for (const section of this.ast.sections ?? []) {
      const found = (section.nodes ?? []).find(n => n.id === previousNodeId);
      if (found) { prev = found; break; }
    }
    if (!prev) {
      throw new Error(
        `Node ${previousNodeId} not found for list continuation. ` +
        `Re-read the document with inspect_document and retry with valid node ids.`
      );
    }
    try {
      const numPr = prev.xml.getElementsByTagName('w:numPr')[0] as Element | undefined;
      const numIdEl = numPr?.getElementsByTagName('w:numId')[0] as Element | undefined;
      const numId = numIdEl?.getAttribute('w:val');
      if (numId != null && numId !== '') return numId;
    } catch { /* fall through to loud error */ }
    throw new Error(
      `Node ${previousNodeId} carries no list numbering (missing w:numPr/w:numId), so there is no list to continue. ` +
      `Create a list on it first ('create'), then continue from the list item id.`
    );
  }

  /**
   * BX-03: serializes ast.numbering back to a standalone word/numbering.xml
   * payload. Returns null when the AST holds no numbering definitions (so
   * rezip never creates a spurious part for documents without lists).
   * Stored stub elements are re-serialized verbatim; when none carry an
   * ownerDocument-serializable form, minimal schema-shaped definitions are
   * emitted instead so the file always parses.
   */
  public serializeNumbering(): string | null {
    const abstract = (this.ast.numbering?.abstractNum ?? {}) as Record<string, Element>;
    const concrete = (this.ast.numbering?.concreteNum ?? {}) as Record<string, Element>;
    const absIds = Object.keys(abstract);
    const numIds = Object.keys(concrete);
    if (absIds.length === 0 && numIds.length === 0) return null;

    const serializer = new XMLSerializer();
    const parts: string[] = [];
    const serializeEl = (el: Element): string | null => {
      try {
        const s = serializer.serializeToString(el);
        if (s && s.includes('<')) return s;
      } catch { /* fall through to minimal shape */ }
      return null;
    };

    // Abstract definitions first (schema order: abstractNum* then num*).
    const orderedAbs = [...absIds].sort((a, b) => Number(a) - Number(b));
    for (const id of orderedAbs) {
      const s = serializeEl(abstract[id]);
      if (s) {
        parts.push(s);
      } else {
        // Fail-closed: when the stored abstract cannot be serialized and its
        // numFmt cannot be read (missing, empty, or throwing), do not guess
        // bullet — guessing would silently demote a numbered list to bullets.
        // Readable numFmt values keep the preserve-type behavior below.
        let fmt: 'bullet' | 'decimal';
        try {
          const fmts = (abstract[id] as Element).getElementsByTagName?.('w:numFmt') ?? [];
          const v = (fmts[0] as Element | undefined)?.getAttribute?.('w:val');
          if (v === 'bullet') fmt = 'bullet';
          else if (v && v !== 'bullet') fmt = 'decimal';
          else {
            throw new Error(
              `Cannot serialize numbering abstract ${id}: w:numFmt is missing or empty, ` +
              `so the list type cannot be determined without risking bullet/number demotion. ` +
              `Re-read the document with inspect_document and retry the save.`
            );
          }
        } catch (err) {
          if (err instanceof Error && err.message.includes(`numbering abstract ${id}`)) throw err;
          throw new Error(
            `Cannot serialize numbering abstract ${id}: w:numFmt is unreadable, ` +
            `so the list type cannot be determined without risking bullet/number demotion. ` +
            `Re-read the document with inspect_document and retry the save.`
          );
        }
        const kind: ListKind = fmt === 'bullet' ? 'bullet' : 'numbered';
        const levels = Array.from({ length: LIST_LEVEL_COUNT }, (_, ilvl) => listLevelXml(kind, ilvl)).join('');
        parts.push(
          `<w:abstractNum w:abstractNumId="${id}">` +
          `<w:nsid w:val="${listNsid(kind, id, new Set())}"/><w:multiLevelType w:val="hybridMultilevel"/>` +
          levels +
          `</w:abstractNum>`
        );
      }
    }
    const orderedNum = [...numIds].sort((a, b) => Number(a) - Number(b));
    for (const id of orderedNum) {
      const s = serializeEl(concrete[id]);
      if (s) {
        parts.push(s);
      } else {
        const absRef = orderedAbs.length > 0 ? orderedAbs[0] : '0';
        parts.push(
          `<w:num w:numId="${id}"><w:abstractNumId w:val="${absRef}"/></w:num>`
        );
      }
    }

    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:numbering xmlns:w="${NUMBERING_NS}">` +
      parts.join('') +
      `</w:numbering>`
    );
  }
}
