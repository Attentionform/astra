import type { Node, PageSetupProperties, Section } from '../ast';

/**
 * The sections a file defines, read from its XML.
 *
 * A section ends at a paragraph that carries a w:sectPr directly in its w:pPr,
 * and the last section ends at the w:sectPr that is a child of w:body. The
 * properties in a w:sectPr describe the section that ENDS with that paragraph.
 * A w:sectPr anywhere else (inside a w:sectPrChange, which keeps the properties
 * from before a tracked change, or inside a table cell, which Word ignores) is
 * not a section.
 */

/** Child elements of `el` with this tag; unlike getElementsByTagName it never reaches into nested elements. */
function directChildren(el: Element, tag: string): Element[] {
  const out: Element[] = [];
  for (let node = el.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && (node as Element).tagName === tag) out.push(node as Element);
  }
  return out;
}

export interface FileSection {
  /** The section's w:sectPr; null only for a last section the file gives no properties. */
  sectPr: Element | null;
  /** The paragraph that ends the section; null for the last section. */
  endParagraph: Element | null;
}

/** The w:sectPr that is a direct child of w:body (the last section's properties), if any. */
export function bodySectPr(body: Element | null | undefined): Element | null {
  return body ? (directChildren(body, 'w:sectPr')[0] ?? null) : null;
}

/** The w:sectPr a paragraph carries in its w:pPr, if any. */
function paragraphSectPr(paragraph: Element): Element | null {
  const pPr = directChildren(paragraph, 'w:pPr')[0];
  return pPr ? (directChildren(pPr, 'w:sectPr')[0] ?? null) : null;
}

/**
 * Every section of the file, in order. Always at least one: a file with no
 * section breaks has the one section that w:body ends. Blocks wrapped in a
 * w:sdt content control count, as they do when node ids are assigned; table
 * cells are not entered. Never throws on an odd shape: what it cannot place as
 * a break it leaves out.
 */
export function listFileSections(body: Element | null | undefined): FileSection[] {
  const sections: FileSection[] = [];
  if (!body) return [{ sectPr: null, endParagraph: null }];

  const visit = (container: Element): void => {
    for (let node = container.firstChild; node; node = node.nextSibling) {
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      if (el.tagName === 'w:p') {
        const sectPr = paragraphSectPr(el);
        if (sectPr) sections.push({ sectPr, endParagraph: el });
      } else if (el.tagName === 'w:sdt') {
        for (const content of directChildren(el, 'w:sdtContent')) visit(content);
      }
    }
  };
  visit(body);

  sections.push({ sectPr: bodySectPr(body), endParagraph: null });
  return sections;
}

/** The page size, orientation and margins one w:sectPr gives its section. */
export function pageSetupFromSectPr(sectPr: Element | null): PageSetupProperties {
  let resolved: PageSetupProperties = {
    orientation: 'portrait',
    paperSize: 'Letter'
  };
  if (!sectPr) return resolved;

  const pgSz = sectPr.getElementsByTagName('w:pgSz')[0];
  if (pgSz) {
    const wVal = parseInt(pgSz.getAttribute('w:w') || '0', 10);
    const hVal = parseInt(pgSz.getAttribute('w:h') || '0', 10);
    const orientAttr = pgSz.getAttribute('w:orient');
    const codeAttr = pgSz.getAttribute('w:code');

    let orientation: 'portrait' | 'landscape' = 'portrait';
    if (orientAttr === 'landscape' || (wVal > hVal && orientAttr !== 'portrait')) {
      orientation = 'landscape';
    }

    let paperSize: 'A4' | 'Letter' = 'Letter';
    const isA4Dimensions =
      (Math.abs(wVal - 11906) < 200 && Math.abs(hVal - 16838) < 200) ||
      (Math.abs(wVal - 16838) < 200 && Math.abs(hVal - 11906) < 200);

    if (codeAttr === '9' || isA4Dimensions) {
      paperSize = 'A4';
    }

    resolved = { ...resolved, orientation, paperSize };
  }

  const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
  if (pgMar) {
    const top = parseInt(pgMar.getAttribute('w:top') || '', 10);
    const bottom = parseInt(pgMar.getAttribute('w:bottom') || '', 10);
    const left = parseInt(pgMar.getAttribute('w:left') || '', 10);
    const right = parseInt(pgMar.getAttribute('w:right') || '', 10);
    if (!isNaN(top)) resolved.marginTop = top;
    if (!isNaN(bottom)) resolved.marginBottom = bottom;
    if (!isNaN(left)) resolved.marginLeft = left;
    if (!isNaN(right)) resolved.marginRight = right;
  }
  return resolved;
}

/**
 * Groups the document's blocks (in order) into the file's sections. Each
 * section ends with its break paragraph, which is the last node it holds; the
 * last section holds whatever follows the final break and may hold nothing. The
 * first section is `sec_root` and the others are `sec_2`, `sec_3`, ... by
 * position. Every node lands in exactly one section, in its original order.
 */
export function splitIntoSections(nodes: Node[], fileSections: FileSection[]): Section[] {
  const indexOfElement = new Map<Element, number>();
  const indexOfId = new Map<string, number>();
  nodes.forEach((node, index) => {
    indexOfElement.set(node.xml as Element, index);
    indexOfId.set(node.id, index);
  });
  const indexOfBreak = (paragraph: Element): number | undefined => {
    const byElement = indexOfElement.get(paragraph);
    if (byElement !== undefined) return byElement;
    // getAttribute gives '' for a missing attribute: no id means no match.
    const id = paragraph.getAttribute('data-astra-id');
    return id ? indexOfId.get(id) : undefined;
  };

  const sections: Section[] = [];
  let start = 0;
  for (const fileSection of fileSections) {
    let end: number;
    if (fileSection.endParagraph === null) {
      end = nodes.length;
    } else {
      const index = indexOfBreak(fileSection.endParagraph);
      // A break that is not one of the document's blocks cannot split anything.
      if (index === undefined || index < start) continue;
      end = index + 1;
    }
    sections.push({
      id: sections.length === 0 ? 'sec_root' : `sec_${sections.length + 1}`,
      heading: null,
      nodes: nodes.slice(start, end),
      pageSetup: pageSetupFromSectPr(fileSection.sectPr),
      version: 1,
      checksum: ''
    });
    start = end;
  }
  return sections;
}

/** True when this block is a paragraph that ends a section, i.e. it carries a w:sectPr directly in its w:pPr. */
export function endsSection(block: Element | null | undefined): boolean {
  return !!block && block.tagName === 'w:p' && paragraphSectPr(block) !== null;
}

/** The section and position of the block with this id, searching every section. */
export function locateNode(sections: Section[], id: string): { section: Section; index: number } | undefined {
  for (const section of sections) {
    const index = section.nodes.findIndex((node) => node.id === id);
    if (index !== -1) return { section, index };
  }
  return undefined;
}

/** The last section: where content with no anchor goes. */
export function lastSection(sections: Section[]): Section | undefined {
  return sections[sections.length - 1];
}

/** The XML document the blocks belong to, taken from the first block found in any section. */
export function ownerDocumentOf(sections: Section[]): Document | null {
  for (const section of sections) {
    for (const node of section.nodes) {
      const owner = (node.xml as Element | undefined)?.ownerDocument;
      if (owner) return owner;
    }
  }
  return null;
}

/**
 * Puts blocks into a section at an index and returns the section that now
 * holds them. A section other than the last ends with its break paragraph, and
 * what follows that paragraph in the file belongs to the next section, so an
 * index past the end of such a section means the start of the next one. For
 * every other index, and for the last section, this is a plain splice.
 */
export function insertIntoSection(sections: Section[], section: Section, index: number, nodes: Node[]): Section {
  const position = sections.indexOf(section);
  const next = position === -1 ? undefined : sections[position + 1];
  const last = section.nodes[section.nodes.length - 1];
  if (next && index >= section.nodes.length && last && endsSection(last.xml as Element)) {
    next.nodes.splice(0, 0, ...nodes);
    return next;
  }
  section.nodes.splice(index, 0, ...nodes);
  return section;
}

/** The first (or, with `fromEnd`, last) element in document order inside `el`, or `el` itself, that is one of the known blocks. */
function knownBlockWithin(el: Element, known: Set<Element>, fromEnd: boolean): Element | null {
  if (known.has(el)) return el;
  for (let child = fromEnd ? el.lastChild : el.firstChild; child; child = fromEnd ? child.previousSibling : child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const hit = knownBlockWithin(child as Element, known, fromEnd);
    if (hit) return hit;
  }
  return null;
}

/**
 * Adds a block that is already in the document to the model, at the position
 * it has in the file: after the nearest earlier block the model holds, else
 * before the nearest later one, else at the end of the document.
 */
export function addNodeInDocumentOrder(sections: Section[], node: Node): void {
  const known = new Set<Element>();
  for (const section of sections) for (const held of section.nodes) known.add(held.xml as Element);

  const find = (element: Element): { section: Section; index: number } | undefined => {
    for (const section of sections) {
      const index = section.nodes.findIndex((held) => held.xml === element);
      if (index !== -1) return { section, index };
    }
    return undefined;
  };

  const start = node.xml as Element;
  for (const fromEnd of [true, false]) {
    for (let current: globalThis.Node | null = start; current && (current as Element).tagName !== 'w:body'; current = current.parentNode) {
      for (let sibling = fromEnd ? current.previousSibling : current.nextSibling; sibling; sibling = fromEnd ? sibling.previousSibling : sibling.nextSibling) {
        if (sibling.nodeType !== 1) continue;
        const hit = knownBlockWithin(sibling as Element, known, fromEnd);
        const at = hit ? find(hit) : undefined;
        if (!at) continue;
        if (fromEnd) insertIntoSection(sections, at.section, at.index + 1, [node]);
        else at.section.nodes.splice(at.index, 0, node);
        return;
      }
    }
  }
  const end = lastSection(sections);
  if (end) end.nodes.push(node);
}

/** True when the model's block and the document's element are the same block: the same element, or the same id. */
function isSameBlock(node: Node | undefined, element: Element | null): boolean {
  if (!node || !element) return false;
  if (node.xml === element) return true;
  const id = element.getAttribute('data-astra-id');
  return !!id && node.id === id;
}

/** The sections already end where the document's section breaks are. */
function sectionsMatchFile(sections: Section[], fileSections: FileSection[]): boolean {
  if (sections.length !== fileSections.length) return false;
  for (let i = 0; i < sections.length - 1; i++) {
    const nodes = sections[i].nodes;
    if (!isSameBlock(nodes[nodes.length - 1], fileSections[i].endParagraph)) return false;
  }
  return true;
}

/**
 * Brings the model's sections in line with the section breaks the document
 * has, in place, and returns a sentence saying so, or null when they already
 * match (the usual case, found with one pass over the document's top level).
 *
 * Operations that add, delete or move a break paragraph change where the file's
 * sections end. The blocks are regrouped at the breaks the file has now, keeping
 * their order. A section keeps the id, version, heading and the settings the
 * file does not carry of the previous section that held its first block. The
 * page size and margins are read from the file again.
 *
 * If a break paragraph is not a block the model holds, nothing is regrouped
 * (splitting there would merge sections that should stay apart) and the
 * sentence says the sections could not be checked.
 */
export function reconcileSections(sections: Section[], body: Element | null | undefined): string | null {
  if (!body || sections.length === 0) return null;
  const fileSections = listFileSections(body);
  if (sectionsMatchFile(sections, fileSections)) return null;

  const blocks = sections.flatMap((section) => section.nodes);
  const known = new Set<Element>(blocks.map((block) => block.xml as Element));
  const knownIds = new Set<string>(blocks.map((block) => block.id));
  const unknownBreak = fileSections.some((fileSection) => {
    const paragraph = fileSection.endParagraph;
    if (paragraph === null || known.has(paragraph)) return false;
    const id = paragraph.getAttribute('data-astra-id');
    return !id || !knownIds.has(id);
  });
  if (unknownBreak) {
    return (
      'The sections could not be checked against the section breaks in the document: a section break is not one of the blocks ' +
      'the document lists. Re-read the document with inspect_document before editing sections.'
    );
  }

  const rebuilt = splitIntoSections(blocks, fileSections);

  // A rebuilt section is the previous section that held its first block: when
  // two sections merge the first one lives on, as in the section.merge operation;
  // when one splits, the part with the original first block keeps it. A section
  // with no blocks takes the previous section at its position if that one is
  // still free.
  const previousOfBlock = new Map<Node, Section>();
  for (const previous of sections) for (const block of previous.nodes) previousOfBlock.set(block, previous);
  const claimed = new Set<Section>();
  const usedIds = new Set<string>();
  const carried = rebuilt.map((section, index): Section | undefined => {
    const candidate = section.nodes.length > 0 ? previousOfBlock.get(section.nodes[0]) : sections[index];
    if (!candidate || claimed.has(candidate) || usedIds.has(candidate.id)) return undefined;
    claimed.add(candidate);
    usedIds.add(candidate.id);
    return candidate;
  });
  rebuilt.forEach((section, index) => {
    const previous = carried[index];
    if (previous) {
      const sameBlocks =
        previous.nodes.length === section.nodes.length && previous.nodes.every((node, i) => node === section.nodes[i]);
      section.id = previous.id;
      section.version = previous.version + (sameBlocks ? 0 : 1);
      section.heading = previous.heading;
      section.checksum = previous.checksum;
      section.pageSetup = { ...(previous.pageSetup ?? {}), ...(section.pageSetup ?? {}) };
    } else {
      let n = index + 1;
      while (usedIds.has(`sec_${n}`)) n++;
      section.id = index === 0 && !usedIds.has('sec_root') ? 'sec_root' : `sec_${n}`;
      usedIds.add(section.id);
    }
  });

  // A document with one section shows it as sec_root, as when it is opened.
  if (rebuilt.length === 1) rebuilt[0].id = 'sec_root';

  const before = sections.length;
  sections.splice(0, sections.length, ...rebuilt);
  return (
    `The sections no longer matched the section breaks in the document (a section break was added, deleted or moved), ` +
    `so they were rebuilt from it: the document now has ${rebuilt.length} section${rebuilt.length === 1 ? '' : 's'}, it had ${before}.` +
    (rebuilt.length < before ? ' Text before a deleted break takes the page setup of the section after it.' : '')
  );
}
