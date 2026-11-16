import { XMLSerializer } from '@xmldom/xmldom';
import { DOMParser } from './dom-parser';
import { relationshipElements } from './relationships-dom';
import { decodePartUri } from './xml-entities';

/**
 * Removes chart and image parts that an edit left unreachable (0.5.3, DN2).
 *
 * Deleting a drawing removes its paragraph but not the relationship, the chart part, its workbook or the
 * picture file. A deleted chart's numbers then stay inside the saved file. This pass removes them.
 *
 * A mistake must cost a skipped cleanup, never a broken or emptied file:
 *  - candidates come from the edit: a relationship is a candidate only if the document used it before the
 *    edit (same id, type and target) or it was created by the edit, and nothing uses it now. A part that was
 *    already unused when the file arrived (left by another tool) is never touched;
 *  - only chart and image relationships of the main document are candidates, and only internal ones;
 *  - "used" is read from every XML part: any prefixed attribute whose value equals the id keeps the
 *    relationship, so a header, a footnote, a VML picture or a tracked change that still points at it wins;
 *  - a part is deleted only if no remaining relationship from any part points at it, and only if its name fits
 *    the shape for its relationship type (a crafted target can never steer a delete elsewhere);
 *  - the whole change is planned on an overlay first and compared with the package as it was: the identities
 *    of relationship targets that point at nothing, and of parts without a content type, must not grow.
 *    The live package is written only after that comparison passes;
 *  - anything that cannot be parsed or looks ambiguous (names that differ only by case) skips that candidate
 *    or the whole pass, with the reason reported.
 */

export interface PackageParts {
  names(): string[];
  read(name: string): Buffer | null;
  write(name: string, data: Buffer): void;
  remove(name: string): void;
}

export interface RelationshipView {
  type: string;
  target: string;
  targetMode?: 'External' | 'Internal';
}

/** What the document used before the edit. Built by `snapshotDrawingUse`. */
export interface DrawingUseBefore {
  /** id -> `type|target` of every chart and image relationship of the main document. */
  relationships: ReadonlyMap<string, string>;
  /** Ids of those relationships that the document referenced. */
  referenced: ReadonlySet<string>;
}

export interface DrawingPruneResult {
  /** Relationship ids removed from word/_rels/document.xml.rels. */
  relationshipIds: string[];
  /** Part names removed. */
  parts: string[];
  /** Why nothing (or something) was not removed, when something looked unsafe. */
  skipped?: string;
  /** Pre-existing defects of the package that the pass saw and left alone. */
  existingDefects: string[];
}


/** Elements by local name: a package may bind the relationships or content-types namespace to a prefix (`<p:Relationship>`). */
function elementsByLocalName(doc: Document, local: string): Element[] {
  return Array.from(doc.getElementsByTagNameNS('*', local));
}



const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const CHART_REL_TYPE = `${REL_NS}/chart`;
export const IMAGE_REL_TYPE = `${REL_NS}/image`;
const DOC_RELS = 'word/_rels/document.xml.rels';
const TYPES_PATH = '[Content_Types].xml';
const MAX_DEPTH = 4;

/**
 * A part is removed with a chart or image when it has the usual name or, for a part Word did not name, the content
 * type of that kind. The name alone says little (a chart may be `word/charts/financial-chart.xml`); the content
 * type, with the relationship that points at it and the proof that nothing else reaches it, says what it is.
 */
const CT = {
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  chartStyle: 'application/vnd.ms-office.chartstyle+xml',
  chartColors: 'application/vnd.ms-office.chartcolorstyle+xml',
  chartShapes: 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml',
  themeOverride: 'application/vnd.openxmlformats-officedocument.themeOverride+xml',
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/** Relationship types of a chart part that go away with it, and the content type each target must have. */
const CHART_CHILD_TYPES: ReadonlyArray<{ type: string; accepts: (name: string, contentType: string | undefined) => boolean }> = [
  { type: `${REL_NS}/package`, accepts: (name, ct) => /^word\/embeddings\/[^/]+$/i.test(name) || ct === CT.workbook },
  { type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle', accepts: (n, ct) => ct === CT.chartStyle || /^word\/charts\/style\d+\.xml$/i.test(n) },
  { type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle', accepts: (n, ct) => ct === CT.chartColors || /^word\/charts\/colors\d+\.xml$/i.test(n) },
  { type: `${REL_NS}/chartUserShapes`, accepts: (n, ct) => ct === CT.chartShapes || /^word\/drawings\/[^/]+\.xml$/i.test(n) },
  { type: `${REL_NS}/themeOverride`, accepts: (n, ct) => ct === CT.themeOverride || /^word\/theme\/themeOverride\d+\.xml$/i.test(n) },
];

type Accepts = (name: string, contentType: string | undefined) => boolean;
const inWordFolder = (name: string): boolean => /^word\/(?!_rels\/).+/i.test(name) && !/\.rels$/i.test(name);
const isChartPart: Accepts = (name, ct) => /^word\/charts\/chart\d+\.xml$/i.test(name) || (inWordFolder(name) && ct === CT.chart);
const isImagePart: Accepts = (name, ct) => /^word\/media\/[^/]+$/i.test(name) || (inWordFolder(name) && (ct ?? '').startsWith('image/'));

/** Content type per part name from `[Content_Types].xml` (override first, then the extension default). */
function contentTypeIndex(typesXml: Buffer | null): (name: string) => string | undefined {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  if (typesXml) {
    const doc = parseStrict(typesXml.toString('utf8'));
    for (const el of elementsByLocalName(doc, 'Override')) overrides.set((el.getAttribute('PartName') ?? '').replace(/^\//, '').toLowerCase(), el.getAttribute('ContentType') ?? '');
    for (const el of elementsByLocalName(doc, 'Default')) defaults.set((el.getAttribute('Extension') ?? '').toLowerCase(), el.getAttribute('ContentType') ?? '');
  }
  return (name) => {
    const lower = name.toLowerCase();
    const override = overrides.get(lower);
    if (override !== undefined) return override;
    const dot = lower.lastIndexOf('.');
    return dot >= 0 ? defaults.get(lower.slice(dot + 1)) : undefined;
  };
}

/** Parses XML or throws on any error: xmldom otherwise repairs broken input. */
const parseStrict = (xml: string): Document => {
  const fail = (message: string) => {
    throw new Error(`unreadable XML: ${message}`);
  };
  const doc = new DOMParser({ errorHandler: { warning: () => undefined, error: fail, fatalError: fail } }).parseFromString(xml, 'text/xml');
  if (!doc?.documentElement || doc.getElementsByTagName('parsererror').length > 0) fail('no root element');
  return doc;
};

const relsPartOf = (part: string): string => {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};

/** Package name a relationship target points at, resolved from the source part's folder; null for anything that is not a plain internal path. */
export function resolvePartName(rawTarget: string, sourcePart: string): string | null {
  const target = decodePartUri(rawTarget);
  if (!target || target.includes('\\') || /^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
  const base = target.startsWith('/') ? [] : sourcePart.split('/').slice(0, -1);
  for (const seg of target.replace(/^\//, '').split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') {
      if (base.length === 0) return null;
      base.pop();
    } else base.push(seg);
  }
  return base.length > 0 ? base.join('/') : null;
}

/** Source part of a rels file (`word/_rels/document.xml.rels` -> `word/document.xml`; `_rels/.rels` -> ''). */
const sourceOfRels = (relsName: string): string => {
  const m = /^(.*?)_rels\/(.*)\.rels$/.exec(relsName);
  return m ? `${m[1]}${m[2]}` : '';
};

/** Every value of a prefixed attribute in the part (a superset of the relationship-namespace attributes on purpose: more ids count as used, so fewer parts go). */
export function prefixedAttributeValues(xml: string, into: Set<string>): void {
  for (const m of xml.matchAll(/[\w.-]+:[\w.-]+\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) into.add(m[1] ?? m[2] ?? '');
}

/** The main document is read from the live model (`documentReferences`): its bytes in the package are only refreshed at save. */
const isUsablePart = (name: string): boolean =>
  /\.(xml|vml)$/i.test(name) && !/\.rels$/i.test(name) && name !== TYPES_PATH && name.toLowerCase() !== 'word/document.xml';

const signatureOf = (rel: RelationshipView): string => `${rel.type}|${rel.target}`;

const isDrawingType = (type: string): boolean => type === CHART_REL_TYPE || type === IMAGE_REL_TYPE;

/**
 * Records which chart and image relationships of the main document exist and which the document references.
 * `referencedIds` is the set of every prefixed attribute value in the document.
 */
export function snapshotDrawingUse(relationships: Record<string, RelationshipView>, referencedIds: ReadonlySet<string>): DrawingUseBefore {
  const rels = new Map<string, string>();
  const referenced = new Set<string>();
  for (const [id, rel] of Object.entries(relationships)) {
    if (!isDrawingType(rel.type) || rel.targetMode === 'External') continue;
    rels.set(id, signatureOf(rel));
    if (referencedIds.has(id)) referenced.add(id);
  }
  return { relationships: rels, referenced };
}

export interface DrawingPruneInput {
  parts: PackageParts;
  /** Relationships of the main document as it will be saved (the package file merged with the model). */
  relationships: Record<string, RelationshipView>;
  before: DrawingUseBefore;
  /** Throws on a hostile XML part (DTD, XInclude...). */
  guard: (xml: string, name: string) => void;
  /** Every prefixed attribute value in the live document. Candidates it holds are dropped before any part is read. */
  documentReferences: () => ReadonlySet<string>;
}

/** A restore after a failed apply failed too: the package may be half edited, so the save must stop. */
export class PruneRestoreError extends Error {}

const NOTHING = (): DrawingPruneResult => ({ relationshipIds: [], parts: [], existingDefects: [] });

interface Overlay {
  read(name: string): Buffer | null;
  names(): string[];
}

/** Relationship targets that point at no part, and parts without a content type, as identities. */
function defectsOf(view: Overlay, docRels: Record<string, RelationshipView>): Set<string> {
  const out = new Set<string>();
  const names = view.names();
  const lower = new Set(names.map((n) => n.toLowerCase()));
  for (const name of names) {
    if (!/\.rels$/i.test(name)) continue;
    const source = sourceOfRels(name);
    if (name === DOC_RELS) continue; // the model view below is authoritative for the main document
    const data = view.read(name);
    if (!data) continue;
    for (const el of Array.from(relationshipElements(parseStrict(data.toString('utf8'))))) {
      if (el.getAttribute('TargetMode') === 'External') continue;
      const where = resolvePartName(el.getAttribute('Target') ?? '', source);
      if (where && !lower.has(where.toLowerCase())) out.add(`missing target: ${name} ${el.getAttribute('Id')} -> ${where}`);
    }
  }
  for (const [id, rel] of Object.entries(docRels)) {
    if (rel.targetMode === 'External') continue;
    const where = resolvePartName(rel.target, 'word/document.xml');
    if (where && !lower.has(where.toLowerCase())) out.add(`missing target: ${DOC_RELS} ${id} -> ${where}`);
  }
  const typesData = view.read(TYPES_PATH);
  if (typesData) {
    const types = parseStrict(typesData.toString('utf8'));
    const defaults = new Set<string>();
    const overrides = new Set<string>();
    for (const el of elementsByLocalName(types, 'Default')) defaults.add((el.getAttribute('Extension') ?? '').toLowerCase());
    for (const el of elementsByLocalName(types, 'Override')) overrides.add((el.getAttribute('PartName') ?? '').replace(/^\//, '').toLowerCase());
    for (const name of names) {
      if (name === TYPES_PATH || name.endsWith('/')) continue;
      const ext = name.includes('.') ? name.slice(name.lastIndexOf('.') + 1).toLowerCase() : '';
      if (!overrides.has(name.toLowerCase()) && !defaults.has(ext)) out.add(`no content type: ${name}`);
    }
  }
  return out;
}

export function pruneUnusedDrawingParts(input: DrawingPruneInput): DrawingPruneResult {
  const { parts, relationships, before, guard } = input;
  const candidates: string[] = [];
  for (const [id, rel] of Object.entries(relationships)) {
    if (!isDrawingType(rel.type) || rel.targetMode === 'External') continue;
    const was = before.relationships.get(id);
    const sig = signatureOf(rel);
    // Present before and unused before with the same target: left by another tool, never touched.
    if (was === sig && !before.referenced.has(id)) continue;
    candidates.push(id);
  }
  if (candidates.length === 0) return NOTHING();

  try {
    // 1. What is used now: the live document first (cheap), then every other XML part of the package.
    const names = parts.names();
    const inDocument = input.documentReferences();
    if (candidates.every((id) => inDocument.has(id))) return NOTHING();
    const used = new Set<string>(inDocument);
    for (const name of names) {
      if (!isUsablePart(name)) continue;
      const data = parts.read(name);
      if (!data) continue;
      const xml = data.toString('utf8');
      guard(xml, name);
      prefixedAttributeValues(xml, used);
    }
    const unused = candidates.filter((id) => !used.has(id));
    if (unused.length === 0) return NOTHING();

    // 2. Plan on an overlay: Map name -> Buffer | null (null = removed).
    const overlay = new Map<string, Buffer | null>();
    const view: Overlay = {
      read: (name) => (overlay.has(name) ? overlay.get(name)! : parts.read(name)),
      names: () => names.filter((n) => !(overlay.has(n) && overlay.get(n) === null)),
    };
    const contentTypeOf = contentTypeIndex(parts.read(TYPES_PATH));
    const lowerCount = new Map<string, number>();
    for (const n of names) lowerCount.set(n.toLowerCase(), (lowerCount.get(n.toLowerCase()) ?? 0) + 1);
    const ambiguous = (n: string) => (lowerCount.get(n.toLowerCase()) ?? 0) > 1;

    // Pre-existing defects, compared after the plan. A file that was already broken is not blamed on us.
    const defectsBefore = defectsOf({ read: (n) => parts.read(n), names: () => names }, relationships);

    // 3. Remove the candidate relationships from the main document's model and, when present, its rels file.
    const remaining: Record<string, RelationshipView> = { ...relationships };
    const removedRelIds: string[] = [];
    const removedTargets: Array<{ target: string; type: string }> = [];
    for (const id of unused) {
      const rel = remaining[id];
      const where = resolvePartName(rel.target, 'word/document.xml');
      if (!where) continue;
      if (ambiguous(where)) continue;
      const accepts = rel.type === CHART_REL_TYPE ? isChartPart : isImagePart;
      if (!accepts(where, contentTypeOf(where))) continue;
      delete remaining[id];
      removedRelIds.push(id);
      removedTargets.push({ target: where, type: rel.type });
    }
    if (removedRelIds.length === 0) return { ...NOTHING(), skipped: 'no candidate had a plain, unambiguous target' };

    const relsData = parts.read(DOC_RELS);
    if (relsData) {
      const doc = parseStrict(relsData.toString('utf8'));
      let changed = false;
      for (const el of Array.from(relationshipElements(doc))) {
        if (removedRelIds.includes(el.getAttribute('Id') ?? '')) {
          el.parentNode?.removeChild(el);
          changed = true;
        }
      }
      if (changed) overlay.set(DOC_RELS, Buffer.from(new XMLSerializer().serializeToString(doc), 'utf8'));
    }

    // 4. Reachability: which names do the remaining relationships of every part still point at.
    const reachable = (): Set<string> => {
      const hit = new Set<string>();
      for (const name of view.names()) {
        if (!/\.rels$/i.test(name) || name === DOC_RELS) continue;
        const data = view.read(name);
        if (!data) continue;
        const source = sourceOfRels(name);
        for (const el of Array.from(relationshipElements(parseStrict(data.toString('utf8'))))) {
          if (el.getAttribute('TargetMode') === 'External') continue;
          const where = resolvePartName(el.getAttribute('Target') ?? '', source);
          if (where) hit.add(where.toLowerCase());
        }
      }
      for (const rel of Object.values(remaining)) {
        if (rel.targetMode === 'External') continue;
        const where = resolvePartName(rel.target, 'word/document.xml');
        if (where) hit.add(where.toLowerCase());
      }
      return hit;
    };

    const removedParts: string[] = [];
    const removePart = (name: string) => {
      if (overlay.get(name) === null || !view.read(name)) return;
      overlay.set(name, null);
      removedParts.push(name);
    };
    const visit = (name: string, accepts: Accepts, depth: number) => {
      if (depth > MAX_DEPTH || !accepts(name, contentTypeOf(name)) || ambiguous(name) || !view.read(name)) return;
      if (reachable().has(name.toLowerCase())) return;
      removePart(name);
      const relsName = relsPartOf(name);
      const childRels = view.read(relsName);
      if (!childRels) return;
      const children: Array<{ where: string; accepts: Accepts }> = [];
      for (const el of Array.from(relationshipElements(parseStrict(childRels.toString('utf8'))))) {
        if (el.getAttribute('TargetMode') === 'External') continue;
        const where = resolvePartName(el.getAttribute('Target') ?? '', name);
        if (!where) continue;
        const type = el.getAttribute('Type') ?? '';
        const child = CHART_CHILD_TYPES.find((c) => c.type === type);
        if (child) children.push({ where, accepts: child.accepts });
        else if (type === IMAGE_REL_TYPE) children.push({ where, accepts: isImagePart });
      }
      removePart(relsName);
      for (const c of children) visit(c.where, c.accepts, depth + 1);
    };
    for (const t of removedTargets) visit(t.target, t.type === CHART_REL_TYPE ? isChartPart : isImagePart, 0);

    // 5. Content-type overrides of the parts that went.
    const typesData = parts.read(TYPES_PATH);
    if (typesData && removedParts.length > 0) {
      const doc = parseStrict(typesData.toString('utf8'));
      const gone = new Set(removedParts.map((n) => n.toLowerCase()));
      let changed = false;
      for (const el of elementsByLocalName(doc, 'Override')) {
        if (gone.has((el.getAttribute('PartName') ?? '').replace(/^\//, '').toLowerCase())) {
          el.parentNode?.removeChild(el);
          changed = true;
        }
      }
      if (changed) overlay.set(TYPES_PATH, Buffer.from(new XMLSerializer().serializeToString(doc), 'utf8'));
    }

    // 6. Compare with the package as it was: nothing new may point at nothing or lack a content type.
    const defectsAfter = defectsOf(view, remaining);
    for (const d of defectsAfter) {
      if (!defectsBefore.has(d)) {
        return { relationshipIds: [], parts: [], skipped: `the cleanup would create a new defect (${d}); nothing was removed`, existingDefects: [...defectsBefore] };
      }
    }

    // 7. Apply. The overlay only holds parts that exist, so the writes cannot fail half way for a missing part.
    const originals = new Map<string, Buffer | null>();
    try {
      for (const [name, data] of overlay) {
        originals.set(name, parts.read(name));
        if (data === null) parts.remove(name);
        else parts.write(name, data);
      }
    } catch (err) {
      try {
        for (const [name, data] of originals) {
          if (data) parts.write(name, data);
          else parts.remove(name);
        }
      } catch (restoreErr) {
        // Never reported as an ordinary skip: the caller aborts the save.
        throw new PruneRestoreError(`package cleanup failed (${(err as Error).message}) and the restore failed (${(restoreErr as Error).message})`);
      }
      throw err;
    }
    return { relationshipIds: removedRelIds, parts: removedParts, existingDefects: [...defectsBefore] };
  } catch (err) {
    if (err instanceof PruneRestoreError) throw err;
    return { ...NOTHING(), skipped: `cleanup skipped: ${(err as Error).message}` };
  }
}
