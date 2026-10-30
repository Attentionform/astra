import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

/**
 * Removes header and footer parts that an edit left with no reference.
 *
 * Deleting a section break drops that section's w:headerReference and
 * w:footerReference elements, but the parts stay in the package: invisible on
 * every page, yet readable by anyone who unzips the file. This pass removes
 * them, with the relationship that pointed at them, their content-type
 * override, their own relationships file, and pictures only they used.
 *
 * It is built so that a mistake costs a skipped cleanup, never a broken file:
 *  - only parts that WERE referenced before the edit and are not any more
 *    (a file that arrives with leftovers from another tool is left alone);
 *  - only plain `headerN.xml` / `footerN.xml` names, so a crafted relationship
 *    target can never steer a delete anywhere else;
 *  - a part stays if its relationship id still appears anywhere in the
 *    document (tracked section-property changes keep old references), if
 *    another relationship or another part's relationships file mentions it;
 *  - anything it cannot parse or does not understand skips the cleanup;
 *  - afterwards it checks the package for references that point at nothing
 *    and restores every touched entry if it created one.
 */

export interface PackageParts {
  names(): string[];
  read(name: string): Buffer | null;
  write(name: string, data: Buffer): void;
  remove(name: string): void;
}

export interface PruneResult {
  /** Relationship ids removed from word/_rels/document.xml.rels. */
  relationshipIds: string[];
  /** Part names removed (`word/header2.xml`, `word/media/image1.png`, ...). */
  parts: string[];
  /** Why nothing was removed, when something looked unsafe or went wrong. */
  skipped?: string;
}

const NOTHING: PruneResult = { relationshipIds: [], parts: [] };
const RELS_PATH = 'word/_rels/document.xml.rels';
const TYPES_PATH = '[Content_Types].xml';
const HEADER_FOOTER_PART = /^(header|footer)\d+\.xml$/;

/** Every r:id on a header or footer reference anywhere in the document, including copies kept by tracked changes. */
export function headerFooterRefIds(doc: Document): Set<string> {
  const ids = new Set<string>();
  for (const tag of ['w:headerReference', 'w:footerReference']) {
    const refs = doc.getElementsByTagName(tag);
    for (let i = 0; i < refs.length; i++) {
      const id = refs[i].getAttribute('r:id');
      if (id) ids.add(id);
    }
  }
  return ids;
}

const text = (data: Buffer | null): string | null => (data ? data.toString('utf8') : null);

/** Parses XML or throws on any error: xmldom otherwise repairs broken input, and a repaired file could be missing entries we would then rewrite. */
const parseStrict = (xml: string): Document => {
  const fail = (message: string) => {
    throw new Error(`unreadable XML: ${message}`);
  };
  const doc = new DOMParser({ errorHandler: { warning: () => undefined, error: fail, fatalError: fail } }).parseFromString(xml, 'text/xml');
  if (!doc?.documentElement || doc.getElementsByTagName('parsererror').length > 0) fail('no root element');
  return doc;
};

const relationshipsOf = (xml: string): Element[] => Array.from(parseStrict(xml).getElementsByTagName('Relationship'));

/** Where an internal relationship target lives, or null when it is not a plain path inside the package. */
const resolveTarget = (target: string, baseDir: string): string | null => {
  if (!target || target.includes('..') || target.includes('\\') || /^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
  return target.startsWith('/') ? target.slice(1) : `${baseDir}${target}`;
};

/**
 * Things the package points at that are not there: relationship targets,
 * r:ids used in the document, and content-type overrides. Compared before and
 * after the pass, so a file that was already broken is not blamed on us.
 */
export function danglingReferences(parts: PackageParts, documentXml: string): Set<string> {
  const out = new Set<string>();
  const present = new Set(parts.names());
  const relsText = text(parts.read(RELS_PATH));
  const ids = new Set<string>();
  if (relsText !== null) {
    for (const rel of relationshipsOf(relsText)) {
      const id = rel.getAttribute('Id');
      if (id) ids.add(id);
      if (rel.getAttribute('TargetMode') === 'External') continue;
      const where = resolveTarget(rel.getAttribute('Target') ?? '', 'word/');
      if (where && !present.has(where)) out.add(`relationship ${id} -> ${where}`);
    }
  }
  for (const match of documentXml.matchAll(/\br:(?:id|embed|link|pict)="([^"]+)"/g)) {
    if (!ids.has(match[1])) out.add(`document uses ${match[1]}`);
  }
  const typesText = text(parts.read(TYPES_PATH));
  if (typesText !== null) {
    const doc = parseStrict(typesText);
    const overrides = doc.getElementsByTagName('Override');
    for (let i = 0; i < overrides.length; i++) {
      const name = (overrides[i].getAttribute('PartName') ?? '').replace(/^\//, '');
      if (name && !present.has(name)) out.add(`override ${name}`);
    }
  }
  return out;
}

export interface PruneInput {
  parts: PackageParts;
  /** The document as it will be saved. */
  documentXml: () => string;
  /** Header/footer relationship ids still used by the document now. */
  referencedNow: ReadonlySet<string>;
  /** Header/footer relationship ids the document used before the edit. */
  referencedBefore: ReadonlySet<string>;
  /** Throws on a hostile XML part (DTD, XInclude...). */
  guard: (xml: string, name: string) => void;
}

export function pruneOrphanedHeaderFooters(input: PruneInput): PruneResult {
  const { parts, referencedNow, referencedBefore } = input;
  const candidates = [...referencedBefore].filter((id) => !referencedNow.has(id));
  if (candidates.length === 0) return NOTHING;

  const touched = new Map<string, Buffer | null>();
  const remember = (name: string) => {
    if (!touched.has(name)) touched.set(name, parts.read(name));
  };
  const restore = () => {
    for (const [name, data] of touched) {
      if (data) parts.write(name, data);
      else parts.remove(name);
    }
  };

  try {
    const documentXml = input.documentXml();
    const relsText = text(parts.read(RELS_PATH));
    if (relsText === null) return { ...NOTHING, skipped: 'no document relationships' };
    input.guard(relsText, RELS_PATH);
    const allRels = relationshipsOf(relsText);
    const names = parts.names();

    // 1. Which candidates are plain header/footer parts that nothing else needs?
    const removeIds = new Set<string>();
    const removeParts = new Map<string, string>(); // part name -> relationship id
    for (const id of candidates) {
      if (documentXml.includes(`"${id}"`) || documentXml.includes(`'${id}'`)) continue;
      const rel = allRels.find((r) => r.getAttribute('Id') === id);
      const type = rel?.getAttribute('Type') ?? '';
      const target = rel?.getAttribute('Target') ?? '';
      if (!rel || rel.getAttribute('TargetMode') === 'External') continue;
      if (!/\/(header|footer)$/.test(type) || !HEADER_FOOTER_PART.test(target)) continue;
      const partName = `word/${target}`;
      if (!names.includes(partName)) continue;
      // another relationship pointing at the same part keeps it alive
      if (allRels.some((r) => r !== rel && r.getAttribute('Target') === target)) continue;
      removeIds.add(id);
      removeParts.set(partName, id);
    }
    if (removeParts.size === 0) return NOTHING;

    const ownRels = (partName: string) => `word/_rels/${partName.slice('word/'.length)}.rels`;
    const dropped = new Set([...removeParts.keys()].map(ownRels));

    // 2. No other relationships file may mention a part we are about to remove.
    const otherRelsFiles = names.filter((n) => n.endsWith('.rels') && n !== RELS_PATH && !dropped.has(n));
    const otherRelsText = new Map<string, string>();
    for (const name of otherRelsFiles) {
      const body = text(parts.read(name)) ?? '';
      input.guard(body, name);
      otherRelsText.set(name, body);
    }
    for (const partName of [...removeParts.keys()]) {
      const file = partName.slice('word/'.length);
      if ([...otherRelsText.values()].some((body) => body.includes(file))) {
        removeIds.delete(removeParts.get(partName) as string);
        removeParts.delete(partName);
      }
    }
    if (removeParts.size === 0) return NOTHING;

    // 3. Pictures that only the removed parts used.
    const removeMedia = new Set<string>();
    for (const partName of removeParts.keys()) {
      const own = text(parts.read(ownRels(partName)));
      if (own === null) continue;
      input.guard(own, ownRels(partName));
      for (const rel of relationshipsOf(own)) {
        if (rel.getAttribute('TargetMode') === 'External') continue;
        const where = resolveTarget(rel.getAttribute('Target') ?? '', 'word/');
        if (!where || !where.startsWith('word/media/') || !names.includes(where)) continue;
        const base = where.slice(where.lastIndexOf('/') + 1);
        const usedElsewhere =
          allRels.some((r) => (r.getAttribute('Target') ?? '').endsWith(base)) ||
          [...otherRelsText.values()].some((body) => body.includes(base)) ||
          documentXml.includes(base);
        if (!usedElsewhere) removeMedia.add(where);
      }
    }

    // 4. Apply, remembering every entry so a failed check can put it all back.
    const danglingBefore = danglingReferences(parts, documentXml);
    remember(RELS_PATH);
    remember(TYPES_PATH);
    for (const partName of removeParts.keys()) {
      remember(partName);
      remember(ownRels(partName));
    }
    for (const name of removeMedia) remember(name);

    const relsDoc = parseStrict(relsText);
    for (const rel of Array.from(relsDoc.getElementsByTagName('Relationship'))) {
      if (removeIds.has(rel.getAttribute('Id') ?? '')) rel.parentNode?.removeChild(rel);
    }
    parts.write(RELS_PATH, Buffer.from(new XMLSerializer().serializeToString(relsDoc), 'utf8'));

    const typesText = text(parts.read(TYPES_PATH));
    if (typesText !== null) {
      input.guard(typesText, TYPES_PATH);
      const typesDoc = parseStrict(typesText);
      for (const override of Array.from(typesDoc.getElementsByTagName('Override'))) {
        const name = (override.getAttribute('PartName') ?? '').replace(/^\//, '');
        if (removeParts.has(name) || removeMedia.has(name)) override.parentNode?.removeChild(override);
      }
      parts.write(TYPES_PATH, Buffer.from(new XMLSerializer().serializeToString(typesDoc), 'utf8'));
    }
    for (const partName of removeParts.keys()) {
      parts.remove(partName);
      if (names.includes(ownRels(partName))) parts.remove(ownRels(partName));
    }
    for (const name of removeMedia) parts.remove(name);

    // 5. Nothing may point at a thing that is not there that did not before.
    const created = [...danglingReferences(parts, documentXml)].filter((item) => !danglingBefore.has(item));
    if (created.length > 0) {
      restore();
      return { ...NOTHING, skipped: `would leave ${created.join(', ')}` };
    }
    return { relationshipIds: [...removeIds], parts: [...removeParts.keys(), ...removeMedia] };
  } catch (error) {
    restore();
    return { ...NOTHING, skipped: error instanceof Error ? error.message : String(error) };
  }
}
