import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node } from '../ast';
import { DocxParser, DocxZipAdapter } from '../parser';
import { BaseOperation } from './base';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { allocateSafeHyperlinkRId } from '../manipulators/relationships';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { SETTINGS_CHILD_ORDER } from '../ooxml/schema-order';

export interface FootnoteOperation extends BaseOperation {
  action: 'footnote';
  operation?: 'insert';
  anchor: string;
  text: string;
  id?: string;
}

export interface EndnoteOperation extends BaseOperation {
  action: 'endnote';
  operation?: 'insert';
  anchor: string;
  text: string;
  id?: string;
}

const SETTINGS_PART = 'word/settings.xml';
const SETTINGS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const SETTINGS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

// Relative order only: unknown tags sort last and keep their positions.
function settingsOrderOf(tag: string): number {
  const idx = SETTINGS_CHILD_ORDER.indexOf(tag);
  return idx === -1 ? SETTINGS_CHILD_ORDER.length : idx;
}

/** Direct element child of the settings root with the given qualified tag. */
function findSettingsChild(root: Element, tag: string): Element | null {
  const all = findAllSettingsChildren(root, tag);
  return all.length > 0 ? all[0] : null;
}

/** Direct element children of the settings root with the given qualified tag. */
function findAllSettingsChildren(root: Element, tag: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < root.childNodes.length; i++) {
    const child = root.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === tag) out.push(child as Element);
  }
  return out;
}

/**
 * Duplicate direct siblings (two <w:footnotePr> under one <w:settings>) are a
 * known Word-repair artifact and invalid OOXML. Normalize to a single child
 * before staging: keep the first direct child (children preserved), drop the
 * rest. Runs on the staged in-memory doc only — the zip is untouched until
 * the write phase commits. A property nested anywhere other than a direct
 * child is genuinely ambiguous and still fails closed in the caller.
 */
function dedupeSettingsChildren(root: Element, tag: string): void {
  const dups = findAllSettingsChildren(root, tag);
  for (let i = 1; i < dups.length; i++) {
    root.removeChild(dups[i]);
  }
}

/**
 * Returns the existing direct child or creates it at the schema-correct
 * position (before the first existing child that sorts after it).
 * An existing child found out of order is moved to its correct slot
 * (same node, children preserved) so footnotePr always precedes endnotePr
 * and both precede compat.
 */
function upsertSettingsChild(root: Element, doc: Document, tag: string): Element {
  const existing = findSettingsChild(root, tag);
  if (existing) {
    // Reposition when out of order: detach then fall through to ordered insert.
    let outOfOrder = false;
    const rank = settingsOrderOf(tag);
    for (let i = 0; i < root.childNodes.length; i++) {
      const child = root.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (child === existing) break;
      if (settingsOrderOf((child as Element).tagName) > rank) {
        outOfOrder = true;
        break;
      }
    }
    // Also out of order when an earlier-ranked sibling follows it.
    if (!outOfOrder) {
      let seen = false;
      for (let i = 0; i < root.childNodes.length; i++) {
        const child = root.childNodes[i];
        if (child.nodeType !== 1) continue;
        if (child === existing) {
          seen = true;
          continue;
        }
        if (seen && settingsOrderOf((child as Element).tagName) < rank) {
          outOfOrder = true;
          break;
        }
      }
    }
    if (!outOfOrder) return existing;
    root.removeChild(existing);
    const target = settingsOrderOf(tag);
    for (let i = 0; i < root.childNodes.length; i++) {
      const child = root.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (settingsOrderOf((child as Element).tagName) > target) {
        root.insertBefore(existing, child);
        return existing;
      }
    }
    root.appendChild(existing);
    return existing;
  }
  const created = createWordElement(doc, tag);
  const target = settingsOrderOf(tag);
  for (let i = 0; i < root.childNodes.length; i++) {
    const child = root.childNodes[i];
    if (child.nodeType !== 1) continue;
    if (settingsOrderOf((child as Element).tagName) > target) {
      root.insertBefore(created, child);
      return created;
    }
  }
  root.appendChild(created);
  return created;
}

/**
 * Fail-closed settings parse: empty input, missing root, parse warnings /
 * errors, and non-w:settings roots all throw with a remedy instead of
 * crashing on undefined or emitting duplicate elements.
 */
function parseSettingsFailClosed(raw: string): Document {
  if (typeof raw !== 'string' || raw.trim() === '') {
    throw new Error(
      `Note settings update failed for ${SETTINGS_PART}: settings part is present but empty. ` +
        `Remedy: delete word/settings.xml from the package (it will be recreated) ` +
        `or restore word/settings.xml from a clean copy of the file, then retry the insert.`
    );
  }
  assertNoXxe(raw, SETTINGS_PART);
  let issue = '';
  const doc = new DOMParser({
    errorHandler: {
      warning(msg: string) {
        if (!issue) issue = String(msg).slice(0, 200);
      },
      error(msg: string) {
        if (!issue) issue = String(msg).slice(0, 200);
      },
      fatalError(msg: string) {
        if (!issue) issue = String(msg).slice(0, 200);
      },
    },
  }).parseFromString(raw, 'text/xml') as unknown as Document | undefined;
  if (issue || !doc || !doc.documentElement) {
    throw new Error(
      `Note settings update failed for ${SETTINGS_PART}: settings XML cannot be parsed safely` +
        (issue ? ` (${issue})` : ' (unparseable)') +
        `. Remedy: restore word/settings.xml from a clean copy of the file, then retry the insert.`
    );
  }
  const root = doc.documentElement as unknown as Element;
  if (root.tagName !== 'w:settings') {
    throw new Error(
      `Note settings update failed for ${SETTINGS_PART}: unexpected root <${root.tagName || '(none)'}> ` +
        `(expected <w:settings>). Remedy: restore word/settings.xml from a clean copy of the file, ` +
        `then retry the insert.`
    );
  }
  return doc;
}

// Note inserts need word/settings.xml to carry both note property elements
// so readers place footnotes and endnotes correctly. Files without a
// settings part (or with one missing these elements) leave endnotes adrift,
// so both the footnote and endnote paths create or patch the part here.
// Loud on failure: the caller gets an error with a fix instead of a
// silently broken package.
function ensureFootnoteEndnoteSettings(zipAdapter: DocxZipAdapter, ast: DocumentAST): void {
  // Atomicity strategy: validate before mutate. Every part update (settings,
  // content types, rels) is parsed and staged in memory first; the zip and
  // ast.relationships are only touched after all stages validate, so a
  // staging failure leaves the package and the AST untouched. The final
  // write phase snapshots originals and restores them if an in-memory write
  // ever throws (not expected to fail, best-effort safety net).
  // Boundary (honest): this helper covers the settings registration only.
  // The caller stages its own note part (footnotes/endnotes XML plus its
  // content-type/rels registration) the same way and appends the
  // document.xml reference run only after every zip write commits, so a
  // staging failure also leaves the anchor paragraph clean. The adapter has
  // no delete, so new-part creation is ordered last: a commit throw restores
  // overwritten parts and the new part never lands half-registered.
  try {
    // Stage 1: settings payload.
    let settingsDoc: Document;
    if (zipAdapter.hasFile(SETTINGS_PART)) {
      settingsDoc = parseSettingsFailClosed(zipAdapter.readText(SETTINGS_PART));
    } else {
      settingsDoc = new DOMParser().parseFromString(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>',
        'text/xml'
      );
    }
    const root = settingsDoc.documentElement as unknown as Element;

    // Duplicate direct siblings normalize to one here (staged doc only);
    // a property nested elsewhere stays a fail-closed corrupt-part error below.
    dedupeSettingsChildren(root, 'w:footnotePr');
    dedupeSettingsChildren(root, 'w:endnotePr');

    // A note property nested somewhere other than a direct settings child
    // signals a corrupt part: adding a direct one would duplicate the
    // element, so fail closed instead of guessing.
    for (const tag of ['w:footnotePr', 'w:endnotePr']) {
      const nested = root.getElementsByTagName(tag);
      if (nested.length > 0 && !findSettingsChild(root, tag)) {
        throw new Error(
          `Note settings update failed for ${SETTINGS_PART}: <${tag}> exists but is not a ` +
            `direct child of <w:settings> (corrupt part). ` +
            `Remedy: restore word/settings.xml from a clean copy of the file, then retry the insert.`
        );
      }
    }

    const footnotePr = upsertSettingsChild(root, settingsDoc, 'w:footnotePr');
    const endnotePr = upsertSettingsChild(root, settingsDoc, 'w:endnotePr');
    // Numbering runs once per document: declare continuous restart for both
    // note kinds when the file does not already choose. Additive only — an
    // explicit eachSect/eachPage choice is preserved. numRestart sorts last
    // in (pos, numFmt, numStart, numRestart), so append keeps schema order.
    for (const pr of [footnotePr, endnotePr]) {
      if (pr.getElementsByTagName('w:numRestart').length === 0) {
        const numRestart = createWordElement(settingsDoc, 'w:numRestart');
        numRestart.setAttribute('w:val', 'continuous');
        pr.appendChild(numRestart);
      }
    }
    // Endnotes collect at the end of the document. pos leads the property
    // children in schema order (pos, numFmt, numStart, numRestart), so insert
    // it before any existing child: appending after a pre-existing numFmt
    // fails schema validation.
    if (endnotePr.getElementsByTagName('w:pos').length === 0) {
      const pos = createWordElement(settingsDoc, 'w:pos');
      pos.setAttribute('w:val', 'docEnd');
      const firstEl = Array.from(endnotePr.childNodes).find(n => n.nodeType === 1) ?? null;
      endnotePr.insertBefore(pos, firstEl);
    }
    const stagedSettings = new XMLSerializer().serializeToString(settingsDoc);
    if (!stagedSettings || !stagedSettings.includes('w:footnotePr') || !stagedSettings.includes('w:endnotePr')) {
      throw new Error(
        `Note settings update failed for ${SETTINGS_PART}: staged payload lost a note property. ` +
          `Remedy: restore word/settings.xml from a clean copy of the file, then retry the insert.`
      );
    }

    // Stage 2: [Content_Types].xml payload (only when registration is missing).
    const ctPath = '[Content_Types].xml';
    let stagedCt: string | null = null;
    if (zipAdapter.hasFile(ctPath)) {
      const ctRaw = zipAdapter.readText(ctPath);
      assertNoXxe(ctRaw, ctPath);
      const ctDoc = new DOMParser().parseFromString(ctRaw, 'text/xml');
      const typesEl = ctDoc.getElementsByTagName('Types')[0];
      if (typesEl) {
        const overrides = typesEl.getElementsByTagName('Override');
        let covered = false;
        for (let i = 0; i < overrides.length; i++) {
          if ((overrides[i] as Element).getAttribute('PartName') === '/word/settings.xml') {
            covered = true;
            break;
          }
        }
        if (!covered) {
          const override = ctDoc.createElement('Override');
          override.setAttribute('PartName', '/word/settings.xml');
          override.setAttribute('ContentType', SETTINGS_CONTENT_TYPE);
          typesEl.appendChild(override);
          stagedCt = new XMLSerializer().serializeToString(ctDoc);
        }
      }
    }

    // Stage 3: document rels payload + reservation (AST untouched until write phase).
    const relsPath = 'word/_rels/document.xml.rels';
    let stagedRels: string | null = null;
    let stagedRId: string | null = null;
    if (zipAdapter.hasFile(relsPath)) {
      const relsRaw = zipAdapter.readText(relsPath);
      assertNoXxe(relsRaw, relsPath);
      const relsDoc = new DOMParser().parseFromString(relsRaw, 'text/xml');
      const relsEl = relsDoc.getElementsByTagName('Relationships')[0];
      if (relsEl) {
        const rels = relsEl.getElementsByTagName('Relationship');
        let covered = false;
        for (let i = 0; i < rels.length; i++) {
          const el = rels[i] as Element;
          if (el.getAttribute('Type') === SETTINGS_REL_TYPE || el.getAttribute('Target') === 'settings.xml') {
            covered = true;
            break;
          }
        }
        if (!covered) {
          stagedRId = allocateSafeHyperlinkRId(ast);
          const rel = relsDoc.createElement('Relationship');
          rel.setAttribute('Id', stagedRId);
          rel.setAttribute('Type', SETTINGS_REL_TYPE);
          rel.setAttribute('Target', 'settings.xml');
          relsEl.appendChild(rel);
          stagedRels = new XMLSerializer().serializeToString(relsDoc);
        }
      }
    }

    // Write phase: all stages validated, now commit.
    const pending: Array<{ path: string; content: string }> = [{ path: SETTINGS_PART, content: stagedSettings }];
    if (stagedCt !== null) pending.push({ path: ctPath, content: stagedCt });
    if (stagedRels !== null) pending.push({ path: relsPath, content: stagedRels });
    const originals = new Map<string, string | null>();
    for (const p of pending) {
      originals.set(p.path, zipAdapter.hasFile(p.path) ? zipAdapter.readText(p.path) : null);
    }
    const hadRelKey = stagedRId !== null && Object.prototype.hasOwnProperty.call(ast.relationships, stagedRId);
    try {
      for (const p of pending) zipAdapter.writeText(p.path, p.content);
      if (stagedRId !== null) {
        ast.relationships[stagedRId] = {
          type: SETTINGS_REL_TYPE,
          target: 'settings.xml',
          targetMode: 'Internal',
        };
      }
    } catch (writeErr) {
      for (const [p, orig] of originals) {
        try {
          if (orig !== null) zipAdapter.writeText(p, orig);
        } catch {
          /* best-effort restore */
        }
      }
      if (stagedRId !== null && !hadRelKey) delete ast.relationships[stagedRId];
      const detail = writeErr instanceof Error ? writeErr.message : String(writeErr);
      throw new Error(
        `Note settings update failed for ${SETTINGS_PART}: commit failed (${detail}); ` +
          `original parts restored where possible. ` +
          `Remedy: restore word/settings.xml, [Content_Types].xml and word/_rels/document.xml.rels ` +
          `from a clean copy of the file, then retry the insert.`
      );
    }
  } catch (e) {
    if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
    if (e instanceof Error && e.message.startsWith('Note settings update failed')) throw e;
    const detail = e instanceof Error ? e.message : String(e);
    throw new Error(
      `Note settings update failed for ${SETTINGS_PART}: ${detail}. ` +
        `Remedy: restore word/settings.xml, [Content_Types].xml and word/_rels/document.xml.rels ` +
        `from a clean copy of the file, then retry the insert.`
    );
  }
}

function ensureNoteStyles(parser: DocxParser | undefined, kind: 'Footnote' | 'Endnote'): void {
  if (!parser) return;
  const stylesDom = parser.getStylesDom();
  if (!stylesDom) return;

  const root = stylesDom.documentElement;
  if (!root) return;

  const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const existingStyles = [
    ...Array.from(root.getElementsByTagName('w:style')),
    ...Array.from(root.getElementsByTagNameNS(wNs, 'style')),
  ];
  const seen = new Set<Element>();
  const existingIds = new Set<string>();
  for (const s of existingStyles) {
    if (seen.has(s)) continue;
    seen.add(s);
    const id = s.getAttribute('w:styleId') || s.getAttribute('styleId') || s.getAttributeNS(wNs, 'styleId') || '';
    if (id) existingIds.add(id);
  }

  const textStyleId = `${kind}Text`;
  const refStyleId = `${kind}Reference`;
  const textNameVal = kind === 'Footnote' ? 'footnote text' : 'endnote text';
  const refNameVal = kind === 'Footnote' ? 'footnote reference' : 'endnote reference';

  if (!existingIds.has(textStyleId)) {
    const styleEl = createWordElement(stylesDom, 'w:style');
    styleEl.setAttribute('w:type', 'paragraph');
    styleEl.setAttribute('w:styleId', textStyleId);

    const nameEl = createWordElement(stylesDom, 'w:name');
    nameEl.setAttribute('w:val', textNameVal);
    styleEl.appendChild(nameEl);

    const basedOnEl = createWordElement(stylesDom, 'w:basedOn');
    basedOnEl.setAttribute('w:val', 'Normal');
    styleEl.appendChild(basedOnEl);

    const uiPriorityEl = createWordElement(stylesDom, 'w:uiPriority');
    uiPriorityEl.setAttribute('w:val', '99');
    styleEl.appendChild(uiPriorityEl);

    const unhideEl = createWordElement(stylesDom, 'w:unhideWhenUsed');
    styleEl.appendChild(unhideEl);

    const pPrEl = createWordElement(stylesDom, 'w:pPr');
    const spacingEl = createWordElement(stylesDom, 'w:spacing');
    spacingEl.setAttribute('w:after', '0');
    spacingEl.setAttribute('w:line', '240');
    spacingEl.setAttribute('w:lineRule', 'auto');
    pPrEl.appendChild(spacingEl);
    styleEl.appendChild(pPrEl);

    const rPrEl = createWordElement(stylesDom, 'w:rPr');
    const szEl = createWordElement(stylesDom, 'w:sz');
    szEl.setAttribute('w:val', '20');
    rPrEl.appendChild(szEl);
    const szCsEl = createWordElement(stylesDom, 'w:szCs');
    szCsEl.setAttribute('w:val', '20');
    rPrEl.appendChild(szCsEl);
    styleEl.appendChild(rPrEl);

    root.appendChild(styleEl);
    existingIds.add(textStyleId);
  }

  if (!existingIds.has(refStyleId)) {
    const styleEl = createWordElement(stylesDom, 'w:style');
    styleEl.setAttribute('w:type', 'character');
    styleEl.setAttribute('w:styleId', refStyleId);

    const nameEl = createWordElement(stylesDom, 'w:name');
    nameEl.setAttribute('w:val', refNameVal);
    styleEl.appendChild(nameEl);

    const uiPriorityEl = createWordElement(stylesDom, 'w:uiPriority');
    uiPriorityEl.setAttribute('w:val', '99');
    styleEl.appendChild(uiPriorityEl);

    const unhideEl = createWordElement(stylesDom, 'w:unhideWhenUsed');
    styleEl.appendChild(unhideEl);

    const rPrEl = createWordElement(stylesDom, 'w:rPr');
    const vertAlignEl = createWordElement(stylesDom, 'w:vertAlign');
    vertAlignEl.setAttribute('w:val', 'superscript');
    rPrEl.appendChild(vertAlignEl);
    styleEl.appendChild(rPrEl);

    root.appendChild(styleEl);
    existingIds.add(refStyleId);
  }
}

export function executeFootnote(op: FootnoteOperation, ast: DocumentAST, parser?: DocxParser): void {
  let targetNode: Node | undefined;
  for (const section of ast.sections) {
    targetNode = section.nodes.find(n => n.id === op.anchor);
    if (targetNode) break;
  }

  if (!targetNode) {
    throw new Error(`Anchor node '${op.anchor}' not found for footnote insertion`);
  }

  const doc = targetNode.xml.ownerDocument;
  if (!doc) throw new Error("Document not found on anchor node");

  // P0-ENG-1: allocate max+1 across BOTH document.xml refs AND footnotes.xml
  // ids. Scanning doc refs only ignores pre-existing footnotes.xml entries
  // (e.g. T8 has zero in-doc refs but ids 1,2 in footnotes.xml) and mints a
  // duplicate w:id=1 → invalid OOXML. Scan both, take max+1.
  const existingFootnoteRefs = doc.getElementsByTagName('w:footnoteReference');
  let nextFootnoteId = 1;
  for (let i = 0; i < existingFootnoteRefs.length; i++) {
    const idVal = parseInt(existingFootnoteRefs[i].getAttribute('w:id') || '0', 10);
    if (idVal >= nextFootnoteId) nextFootnoteId = idVal + 1;
  }
  if (parser) {
    try {
      const zipAdapter = parser.getZipAdapter() as unknown as {
        hasFile: (p: string) => boolean;
        readText: (p: string) => string;
      };
      const footnotesPath = 'word/footnotes.xml';
      if (zipAdapter.hasFile(footnotesPath)) {
        const fnXml = zipAdapter.readText(footnotesPath);
        assertNoXxe(fnXml, footnotesPath);
        const fnDoc: Document = new DOMParser().parseFromString(fnXml, 'text/xml');
        const fnEls = fnDoc.getElementsByTagName('w:footnote');
        for (let i = 0; i < fnEls.length; i++) {
          const raw = (fnEls[i] as Element).getAttribute('w:id');
          if (raw == null || raw === '') continue;
          const n = parseInt(raw, 10);
          // Separator ids (-1, 0) never advance the allocator past 1.
          if (Number.isFinite(n) && n >= nextFootnoteId) nextFootnoteId = n + 1;
        }
      }
    } catch (e) {
      if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
      /* best-effort: doc-ref scan already ran; footnotes.xml scan is additive */
    }
  }

  // Build the superscript reference run detached. It joins the live
  // document DOM only after every zip write below has committed, so a
  // settings or part staging failure leaves the anchor paragraph clean
  // (no partial insert). The live DOM leads the zip bytes and rezip flushes
  // it, so appending earlier would leak a run on failure with no undo.
  const run = createWordElement(doc, 'w:r');
  const rPr = createWordElement(doc, 'w:rPr');
  const rStyle = createWordElement(doc, 'w:rStyle');
  rStyle.setAttribute('w:val', 'FootnoteReference');
  const vertAlign = createWordElement(doc, 'w:vertAlign');
  vertAlign.setAttribute('w:val', 'superscript');
  rPr.appendChild(rStyle);
  rPr.appendChild(vertAlign);
  run.appendChild(rPr);

  const footnoteRef = createWordElement(doc, 'w:footnoteReference');
  footnoteRef.setAttribute('w:id', String(nextFootnoteId));
  run.appendChild(footnoteRef);

  // No package sync without a parser: in-memory DOM append only.
  if (!parser) {
    targetNode.xml.appendChild(run);
    targetNode.version += 1;
    return;
  }

  // Settings first: throws on corrupt parts before any mutation, DOM clean.
  const zipAdapter = parser.getZipAdapter();
  ensureFootnoteEndnoteSettings(zipAdapter, ast);
  const footnotesPath = 'word/footnotes.xml';
  let footnotesDoc: Document;

  if (zipAdapter.hasFile(footnotesPath)) {
    const existingXml = zipAdapter.readText(footnotesPath);
    assertNoXxe(existingXml, footnotesPath);
    footnotesDoc = new DOMParser().parseFromString(existingXml, 'text/xml');
  } else {
    const baseXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
</w:footnotes>`;
    footnotesDoc = new DOMParser().parseFromString(baseXml, 'text/xml');
  }

  // Add new footnote entry to the in-memory doc (still uncommitted).
  const footnotesRoot = footnotesDoc.getElementsByTagName('w:footnotes')[0] || footnotesDoc.documentElement;
  const newFn = createWordElement(footnotesDoc, 'w:footnote');
  newFn.setAttribute('w:id', String(nextFootnoteId));

  const p = createWordElement(footnotesDoc, 'w:p');
  const pPr = createWordElement(footnotesDoc, 'w:pPr');
  const pStyle = createWordElement(footnotesDoc, 'w:pStyle');
  pStyle.setAttribute('w:val', 'FootnoteText');
  pPr.appendChild(pStyle);
  p.appendChild(pPr);

  const refRun = createWordElement(footnotesDoc, 'w:r');
  const refRPr = createWordElement(footnotesDoc, 'w:rPr');
  const refStyle = createWordElement(footnotesDoc, 'w:rStyle');
  refStyle.setAttribute('w:val', 'FootnoteReference');
  const refVert = createWordElement(footnotesDoc, 'w:vertAlign');
  refVert.setAttribute('w:val', 'superscript');
  refRPr.appendChild(refStyle);
  refRPr.appendChild(refVert);
  refRun.appendChild(refRPr);
  const fnRefElem = createWordElement(footnotesDoc, 'w:footnoteRef');
  refRun.appendChild(fnRefElem);
  p.appendChild(refRun);

  const textRun = createWordElement(footnotesDoc, 'w:r');
  const spaceRun = createWordElement(footnotesDoc, 'w:t');
  spaceRun.setAttribute('xml:space', 'preserve');
  spaceRun.textContent = ` ${op.text}`;
  textRun.appendChild(spaceRun);
  p.appendChild(textRun);

  newFn.appendChild(p);
  footnotesRoot.appendChild(newFn);
  const stagedFootnotes = new XMLSerializer().serializeToString(footnotesDoc);
  if (!stagedFootnotes || !stagedFootnotes.includes(`w:id="${nextFootnoteId}"`)) {
    throw new Error(
      `Note settings update failed for ${footnotesPath}: staged payload lost the new entry. ` +
        `Remedy: restore word/footnotes.xml from a clean copy of the file, then retry the insert.`
    );
  }

  // Part registration staged in memory too (skipped when already covered,
  // so repeat inserts never duplicate the Override or Relationship).
  const ctPath = '[Content_Types].xml';
  let stagedCt: string | null = null;
  if (!zipAdapter.hasFile(footnotesPath) && zipAdapter.hasFile(ctPath)) {
    const ctRaw = zipAdapter.readText(ctPath);
    assertNoXxe(ctRaw, ctPath);
    const ctDoc = new DOMParser().parseFromString(ctRaw, 'text/xml');
    const typesEl = ctDoc.getElementsByTagName('Types')[0];
    if (typesEl) {
      const overrides = typesEl.getElementsByTagName('Override');
      let covered = false;
      for (let i = 0; i < overrides.length; i++) {
        if ((overrides[i] as Element).getAttribute('PartName') === '/word/footnotes.xml') {
          covered = true;
          break;
        }
      }
      if (!covered) {
        const override = ctDoc.createElement('Override');
        override.setAttribute('PartName', '/word/footnotes.xml');
        override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
        typesEl.appendChild(override);
        stagedCt = new XMLSerializer().serializeToString(ctDoc);
      }
    }
  }

  const relsPath = 'word/_rels/document.xml.rels';
  let stagedRels: string | null = null;
  let stagedRId: string | null = null;
  if (!zipAdapter.hasFile(footnotesPath) && zipAdapter.hasFile(relsPath)) {
    const relsRaw = zipAdapter.readText(relsPath);
    assertNoXxe(relsRaw, relsPath);
    const relsDoc = new DOMParser().parseFromString(relsRaw, 'text/xml');
    const relsEl = relsDoc.getElementsByTagName('Relationships')[0];
    if (relsEl) {
      const rels = relsEl.getElementsByTagName('Relationship');
      let covered = false;
      for (let i = 0; i < rels.length; i++) {
        const el = rels[i] as Element;
        if (
          el.getAttribute('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes' ||
          el.getAttribute('Target') === 'footnotes.xml'
        ) {
          covered = true;
          break;
        }
      }
      if (!covered) {
        // Collision-safe rId: reuse the shared atomic allocator (canonical
        // rIdN + process-wide reservation across callers). Reserved here but
        // published to ast.relationships only on commit, so a failed commit
        // leaves the AST untouched.
        stagedRId = allocateSafeHyperlinkRId(ast);
        const rel = relsDoc.createElement('Relationship');
        rel.setAttribute('Id', stagedRId);
        rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes');
        rel.setAttribute('Target', 'footnotes.xml');
        relsEl.appendChild(rel);
        stagedRels = new XMLSerializer().serializeToString(relsDoc);
      }
    }
  }

  // Commit: payload last, so a throw restores the overwritten registrations
  // and the new part never lands half-registered. The adapter has no delete,
  // so ordering (overwrites first, new part last) is the safety mechanism;
  // a failure here leaves previously existing parts byte-identical.
  const pending: Array<{ path: string; content: string }> = [];
  if (stagedCt !== null) pending.push({ path: ctPath, content: stagedCt });
  if (stagedRels !== null) pending.push({ path: relsPath, content: stagedRels });
  pending.push({ path: footnotesPath, content: stagedFootnotes });
  const originals = new Map<string, string | null>();
  for (const item of pending) {
    originals.set(item.path, zipAdapter.hasFile(item.path) ? zipAdapter.readText(item.path) : null);
  }
  const hadRelKey = stagedRId !== null && Object.prototype.hasOwnProperty.call(ast.relationships, stagedRId);
  try {
    for (const item of pending) zipAdapter.writeText(item.path, item.content);
    if (stagedRId !== null) {
      ast.relationships[stagedRId] = {
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
        target: 'footnotes.xml',
        targetMode: 'Internal',
      };
    }
  } catch (writeErr) {
    for (const [entryPath, orig] of originals) {
      try {
        if (orig !== null) zipAdapter.writeText(entryPath, orig);
      } catch {
        /* best-effort restore */
      }
    }
    if (stagedRId !== null && !hadRelKey) delete ast.relationships[stagedRId];
    const detail = writeErr instanceof Error ? writeErr.message : String(writeErr);
    throw new Error(
      `Note settings update failed for ${footnotesPath}: commit failed (${detail}); ` +
        `original parts restored where possible. ` +
        `Remedy: restore word/footnotes.xml, [Content_Types].xml and word/_rels/document.xml.rels ` +
        `from a clean copy of the file, then retry the insert.`
    );
  }

  // All zip writes committed: now join the live DOM.
  ensureNoteStyles(parser, 'Footnote');
  targetNode.xml.appendChild(run);
  targetNode.version += 1;
}

export function executeEndnote(op: EndnoteOperation, ast: DocumentAST, parser?: DocxParser): void {
  let targetNode: Node | undefined;
  for (const section of ast.sections) {
    targetNode = section.nodes.find(n => n.id === op.anchor);
    if (targetNode) break;
  }

  if (!targetNode) {
    throw new Error(`Anchor node '${op.anchor}' not found for endnote insertion`);
  }

  const doc = targetNode.xml.ownerDocument;
  if (!doc) throw new Error("Document not found on anchor node");

  // Mirror of the footnote max+1 allocator: scan BOTH document.xml
  // w:endnoteReference ids AND word/endnotes.xml w:endnote ids so a doc
  // with zero in-doc refs but pre-existing endnotes (e.g. T8 ids 1,2)
  // never re-mints a duplicate w:id → invalid OOXML.
  const existingEndnoteRefs = doc.getElementsByTagName('w:endnoteReference');
  let nextEndnoteId = 1;
  for (let i = 0; i < existingEndnoteRefs.length; i++) {
    const idVal = parseInt(existingEndnoteRefs[i].getAttribute('w:id') || '0', 10);
    if (idVal >= nextEndnoteId) nextEndnoteId = idVal + 1;
  }
  if (parser) {
    try {
      const zipAdapter = parser.getZipAdapter() as unknown as {
        hasFile: (p: string) => boolean;
        readText: (p: string) => string;
      };
      const endnotesPath = 'word/endnotes.xml';
      if (zipAdapter.hasFile(endnotesPath)) {
        const enXml = zipAdapter.readText(endnotesPath);
        assertNoXxe(enXml, endnotesPath);
        const enDoc: Document = new DOMParser().parseFromString(enXml, 'text/xml');
        const enEls = enDoc.getElementsByTagName('w:endnote');
        for (let i = 0; i < enEls.length; i++) {
          const raw = (enEls[i] as Element).getAttribute('w:id');
          if (raw == null || raw === '') continue;
          const n = parseInt(raw, 10);
          // Separator ids (-1, 0) never advance the allocator past 1.
          if (Number.isFinite(n) && n >= nextEndnoteId) nextEndnoteId = n + 1;
        }
      }
    } catch (e) {
      if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
      /* best-effort: doc-ref scan already ran; endnotes.xml scan is additive */
    }
  }

  // Build the superscript reference run detached. It joins the live
  // document DOM only after every zip write below has committed, so a
  // settings or part staging failure leaves the anchor paragraph clean
  // (no partial insert). The live DOM leads the zip bytes and rezip flushes
  // it, so appending earlier would leak a run on failure with no undo.
  const run = createWordElement(doc, 'w:r');
  const rPr = createWordElement(doc, 'w:rPr');
  const rStyle = createWordElement(doc, 'w:rStyle');
  rStyle.setAttribute('w:val', 'EndnoteReference');
  const vertAlign = createWordElement(doc, 'w:vertAlign');
  vertAlign.setAttribute('w:val', 'superscript');
  rPr.appendChild(rStyle);
  rPr.appendChild(vertAlign);
  run.appendChild(rPr);

  const endnoteRef = createWordElement(doc, 'w:endnoteReference');
  endnoteRef.setAttribute('w:id', String(nextEndnoteId));
  run.appendChild(endnoteRef);

  // No package sync without a parser: in-memory DOM append only.
  if (!parser) {
    targetNode.xml.appendChild(run);
    targetNode.version += 1;
    return;
  }

  // Settings first: throws on corrupt parts before any mutation, DOM clean.
  const zipAdapter = parser.getZipAdapter();
  ensureFootnoteEndnoteSettings(zipAdapter, ast);
  const endnotesPath = 'word/endnotes.xml';
  let endnotesDoc: Document;

  if (zipAdapter.hasFile(endnotesPath)) {
    const existingXml = zipAdapter.readText(endnotesPath);
    assertNoXxe(existingXml, endnotesPath);
    endnotesDoc = new DOMParser().parseFromString(existingXml, 'text/xml');
  } else {
    const baseXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
</w:endnotes>`;
    endnotesDoc = new DOMParser().parseFromString(baseXml, 'text/xml');
  }

  // Add new endnote entry to the in-memory doc (still uncommitted).
  const endnotesRoot = endnotesDoc.getElementsByTagName('w:endnotes')[0] || endnotesDoc.documentElement;
  const newEn = createWordElement(endnotesDoc, 'w:endnote');
  newEn.setAttribute('w:id', String(nextEndnoteId));

  const p = createWordElement(endnotesDoc, 'w:p');
  const pPr = createWordElement(endnotesDoc, 'w:pPr');
  const pStyle = createWordElement(endnotesDoc, 'w:pStyle');
  pStyle.setAttribute('w:val', 'EndnoteText');
  pPr.appendChild(pStyle);
  p.appendChild(pPr);

  const refRun = createWordElement(endnotesDoc, 'w:r');
  const refRPr = createWordElement(endnotesDoc, 'w:rPr');
  const refStyle = createWordElement(endnotesDoc, 'w:rStyle');
  refStyle.setAttribute('w:val', 'EndnoteReference');
  const refVert = createWordElement(endnotesDoc, 'w:vertAlign');
  refVert.setAttribute('w:val', 'superscript');
  refRPr.appendChild(refStyle);
  refRPr.appendChild(refVert);
  refRun.appendChild(refRPr);
  const enRefElem = createWordElement(endnotesDoc, 'w:endnoteRef');
  refRun.appendChild(enRefElem);
  p.appendChild(refRun);

  const textRun = createWordElement(endnotesDoc, 'w:r');
  const spaceRun = createWordElement(endnotesDoc, 'w:t');
  spaceRun.setAttribute('xml:space', 'preserve');
  spaceRun.textContent = ` ${op.text}`;
  textRun.appendChild(spaceRun);
  p.appendChild(textRun);

  newEn.appendChild(p);
  endnotesRoot.appendChild(newEn);
  const stagedEndnotes = new XMLSerializer().serializeToString(endnotesDoc);
  if (!stagedEndnotes || !stagedEndnotes.includes(`w:id="${nextEndnoteId}"`)) {
    throw new Error(
      `Note settings update failed for ${endnotesPath}: staged payload lost the new entry. ` +
        `Remedy: restore word/endnotes.xml from a clean copy of the file, then retry the insert.`
    );
  }

  // Part registration staged in memory too (skipped when already covered,
  // so repeat inserts never duplicate the Override or Relationship).
  const ctPath = '[Content_Types].xml';
  let stagedCt: string | null = null;
  if (!zipAdapter.hasFile(endnotesPath) && zipAdapter.hasFile(ctPath)) {
    const ctRaw = zipAdapter.readText(ctPath);
    assertNoXxe(ctRaw, ctPath);
    const ctDoc = new DOMParser().parseFromString(ctRaw, 'text/xml');
    const typesEl = ctDoc.getElementsByTagName('Types')[0];
    if (typesEl) {
      const overrides = typesEl.getElementsByTagName('Override');
      let covered = false;
      for (let i = 0; i < overrides.length; i++) {
        if ((overrides[i] as Element).getAttribute('PartName') === '/word/endnotes.xml') {
          covered = true;
          break;
        }
      }
      if (!covered) {
        const override = ctDoc.createElement('Override');
        override.setAttribute('PartName', '/word/endnotes.xml');
        override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml');
        typesEl.appendChild(override);
        stagedCt = new XMLSerializer().serializeToString(ctDoc);
      }
    }
  }

  const relsPath = 'word/_rels/document.xml.rels';
  let stagedRels: string | null = null;
  let stagedRId: string | null = null;
  if (!zipAdapter.hasFile(endnotesPath) && zipAdapter.hasFile(relsPath)) {
    const relsRaw = zipAdapter.readText(relsPath);
    assertNoXxe(relsRaw, relsPath);
    const relsDoc = new DOMParser().parseFromString(relsRaw, 'text/xml');
    const relsEl = relsDoc.getElementsByTagName('Relationships')[0];
    if (relsEl) {
      const rels = relsEl.getElementsByTagName('Relationship');
      let covered = false;
      for (let i = 0; i < rels.length; i++) {
        const el = rels[i] as Element;
        if (
          el.getAttribute('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes' ||
          el.getAttribute('Target') === 'endnotes.xml'
        ) {
          covered = true;
          break;
        }
      }
      if (!covered) {
        // Collision-safe rId: reuse the shared atomic allocator (canonical
        // rIdN + process-wide reservation across callers). Reserved here but
        // published to ast.relationships only on commit, so a failed commit
        // leaves the AST untouched.
        stagedRId = allocateSafeHyperlinkRId(ast);
        const rel = relsDoc.createElement('Relationship');
        rel.setAttribute('Id', stagedRId);
        rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes');
        rel.setAttribute('Target', 'endnotes.xml');
        relsEl.appendChild(rel);
        stagedRels = new XMLSerializer().serializeToString(relsDoc);
      }
    }
  }

  // Commit: payload last, so a throw restores the overwritten registrations
  // and the new part never lands half-registered. The adapter has no delete,
  // so ordering (overwrites first, new part last) is the safety mechanism;
  // a failure here leaves previously existing parts byte-identical.
  const pending: Array<{ path: string; content: string }> = [];
  if (stagedCt !== null) pending.push({ path: ctPath, content: stagedCt });
  if (stagedRels !== null) pending.push({ path: relsPath, content: stagedRels });
  pending.push({ path: endnotesPath, content: stagedEndnotes });
  const originals = new Map<string, string | null>();
  for (const item of pending) {
    originals.set(item.path, zipAdapter.hasFile(item.path) ? zipAdapter.readText(item.path) : null);
  }
  const hadRelKey = stagedRId !== null && Object.prototype.hasOwnProperty.call(ast.relationships, stagedRId);
  try {
    for (const item of pending) zipAdapter.writeText(item.path, item.content);
    if (stagedRId !== null) {
      ast.relationships[stagedRId] = {
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
        target: 'endnotes.xml',
        targetMode: 'Internal',
      };
    }
  } catch (writeErr) {
    for (const [entryPath, orig] of originals) {
      try {
        if (orig !== null) zipAdapter.writeText(entryPath, orig);
      } catch {
        /* best-effort restore */
      }
    }
    if (stagedRId !== null && !hadRelKey) delete ast.relationships[stagedRId];
    const detail = writeErr instanceof Error ? writeErr.message : String(writeErr);
    throw new Error(
      `Note settings update failed for ${endnotesPath}: commit failed (${detail}); ` +
        `original parts restored where possible. ` +
        `Remedy: restore word/endnotes.xml, [Content_Types].xml and word/_rels/document.xml.rels ` +
        `from a clean copy of the file, then retry the insert.`
    );
  }

  // All zip writes committed: now join the live DOM.
  ensureNoteStyles(parser, 'Endnote');
  targetNode.xml.appendChild(run);
  targetNode.version += 1;
}
