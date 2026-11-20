import type { DocxParser } from '../parser';
import { REGION_END, REGION_START, findRegion, findSlot, hashNodes } from './template-canonical';
import { TEMPLATE_MANIFEST } from './template-manifest.generated';
import { unescapeMarkedText } from './text-spans';

/**
 * What the compose shortcut may do with a document that came from a built-in template.
 *
 * A template carries a custom property `AstraTemplate` (`<name>@<design version>`) and two empty bookmarks that bound its
 * placeholder body. Compose replaces that body only when the shipped manifest says the body is exactly the original. Any doubt
 * gives `append`, which deletes nothing, with a reason code. The document is never trusted to describe itself: the expected
 * hashes come from the manifest that ships with the engine.
 */
export type TemplateReason =
  | 'TEMPLATE_MARKER_INVALID'
  | 'TEMPLATE_MANIFEST_MISSING'
  | 'TEMPLATE_BOOKMARKS_MISSING'
  | 'TEMPLATE_REGION_EDITED'
  | 'TEMPLATE_ALREADY_COMPOSED';

export type TemplateSeedPlan =
  | { kind: 'ordinary' }
  | { kind: 'append'; template: string; reason: TemplateReason; message: string }
  | {
      kind: 'replace';
      template: string;
      /** AST node ids of the body to remove, in document order. */
      regionIds: string[];
      /** Composed content goes after this node, inside the region. */
      anchorId: string;
      /** The boundary bookmarks and the slot, for the cleanup after a successful compose. */
      boundaries: Element[];
      slot: { paragraph: Element; start: Element; end: Element; nodes: Node[] } | null;
    };

export const MARKER_PART = 'docProps/custom.xml';
const MARKER_RE = /<property\b[^>]*\bname="AstraTemplate"[^>]*>\s*<vt:lpwstr>([^<]*)<\/vt:lpwstr>\s*<\/property>/;

export interface TemplateMarker {
  template: string;
  version: number;
  composed: boolean;
}

export function readTemplateMarker(parser: DocxParser): TemplateMarker | null | 'invalid' {
  const xml = parser.getZipAdapter().readTextOrNull(MARKER_PART);
  if (xml === null) return null;
  const m = MARKER_RE.exec(xml);
  if (!m) return null;
  const parsed = /^([a-z0-9-]+)@(\d+)(:composed)?$/.exec(m[1].trim());
  if (!parsed) return 'invalid';
  return { template: parsed[1], version: Number(parsed[2]), composed: parsed[3] !== undefined };
}

/** Rewrites the marker to the composed state, so a second compose appends. */
export function markTemplateComposed(parser: DocxParser, marker: TemplateMarker): void {
  const zip = parser.getZipAdapter();
  const xml = zip.readTextOrNull(MARKER_PART);
  if (xml === null) return;
  zip.writeText(MARKER_PART, xml.replace(MARKER_RE, (whole, value: string) => whole.replace(value, `${marker.template}@${marker.version}:composed`)));
}

export function planTemplateSeed(parser: DocxParser): TemplateSeedPlan {
  const marker = readTemplateMarker(parser);
  if (marker === null) return { kind: 'ordinary' };
  if (marker === 'invalid') {
    return { kind: 'append', template: '', reason: 'TEMPLATE_MARKER_INVALID', message: 'The AstraTemplate property is not in the form <name>@<version>, so the document is treated as an ordinary document.' };
  }
  const { template } = marker;
  // The blank template has no placeholder body: a lone empty paragraph is recognised on its own (pristine-seed.ts).
  if (template === 'blank') return { kind: 'ordinary' };
  if (marker.composed) {
    return { kind: 'append', template, reason: 'TEMPLATE_ALREADY_COMPOSED', message: `The ${template} template already received composed content, so new content is appended.` };
  }
  const entry = TEMPLATE_MANIFEST[`${template}@${marker.version}`];
  if (!entry) {
    return { kind: 'append', template, reason: 'TEMPLATE_MANIFEST_MISSING', message: `This engine has no manifest for ${template} design version ${marker.version}, so the document is treated as an ordinary document.` };
  }
  const body = parser.getDocumentDom().getElementsByTagName('w:body')[0];
  const region = body ? findRegion(body) : null;
  if (!body || !region) {
    return { kind: 'append', template, reason: 'TEMPLATE_BOOKMARKS_MISSING', message: `The ${REGION_START} and ${REGION_END} bookmarks are missing, duplicated or out of order, so nothing is replaced and the new content is appended.` };
  }
  if (!entry.region.includes(hashNodes(region.nodes))) {
    return { kind: 'append', template, reason: 'TEMPLATE_REGION_EDITED', message: `The placeholder body of the ${template} template was edited, so nothing is deleted and the new content is appended.` };
  }
  const ids: string[] = [];
  const idOf = new Map<Element, string>();
  for (const [id, el] of parser.nodeMap) idOf.set(el, id);
  for (const el of region.nodes) {
    const id = idOf.get(el);
    if (id === undefined) {
      return { kind: 'append', template, reason: 'TEMPLATE_REGION_EDITED', message: `The placeholder body of the ${template} template holds content that cannot be addressed, so nothing is deleted and the new content is appended.` };
    }
    ids.push(id);
  }
  if (ids.length === 0) {
    return { kind: 'append', template, reason: 'TEMPLATE_REGION_EDITED', message: `The placeholder body of the ${template} template is empty, so the new content is appended.` };
  }
  let slot: ReturnType<typeof findSlot> = null;
  const slotHash = entry.slots.title;
  if (slotHash !== undefined) {
    const found = findSlot(body, 'title');
    if (found && hashNodes(found.nodes) === slotHash) slot = found;
  }
  return { kind: 'replace', template, regionIds: ids, anchorId: ids[ids.length - 1], boundaries: [region.start, region.end], slot };
}

/**
 * The text of a leading level-1 heading that may fill the template's title slot, with the Markdown that remains.
 * Only a plain-text heading qualifies: inline formatting would end up in the slot as raw marks.
 */
export function splitLeadingTitle(markdown: string): { title: string; rest: string } | null {
  const m = /^(?:[ \t]*\r?\n)*#[ \t]+([^\r\n]+?)[ \t]*#*[ \t]*(?:\r?\n|$)/.exec(markdown);
  if (!m) return null;
  const title = m[1].trim();
  if (title === '' || /[*_`[\]<>\\|]/.test(title)) return null;
  // The compiler reads entities such as `&amp;` in a heading; the title slot must read them the same way.
  return { title: unescapeMarkedText(title), rest: markdown.slice(m[0].length) };
}

/** Replaces the placeholder run(s) of the slot with plain text, keeping any direct run formatting such as bold. */
export function fillSlot(slot: NonNullable<Extract<TemplateSeedPlan, { kind: 'replace' }>['slot']>, text: string): void {
  const doc = slot.paragraph.ownerDocument;
  const firstRun = slot.nodes.find((n): n is Element => n.nodeType === 1 && (n as Element).tagName === 'w:r');
  const run = doc.createElementNS(slot.paragraph.namespaceURI, 'w:r');
  const rPr = firstRun ? Array.from(firstRun.childNodes).find((c) => c.nodeType === 1 && (c as Element).tagName === 'w:rPr') : undefined;
  if (rPr) {
    const copy = rPr.cloneNode(true) as Element;
    for (const style of Array.from(copy.getElementsByTagName('w:rStyle'))) copy.removeChild(style);
    if (copy.childNodes.length > 0) run.appendChild(copy);
  }
  const t = doc.createElementNS(slot.paragraph.namespaceURI, 'w:t');
  t.setAttribute('xml:space', 'preserve');
  t.appendChild(doc.createTextNode(text));
  run.appendChild(t);
  for (const n of slot.nodes) slot.paragraph.removeChild(n);
  slot.paragraph.insertBefore(run, slot.end);
}

/** Removes the two boundary bookmarks (start and end of each). */
export function removeBoundaries(plan: Extract<TemplateSeedPlan, { kind: 'replace' }>): void {
  for (const start of plan.boundaries) {
    const end = start.nextSibling;
    const parent = start.parentNode;
    if (!parent) continue;
    if (end && end.nodeType === 1 && (end as Element).tagName === 'w:bookmarkEnd') parent.removeChild(end);
    parent.removeChild(start);
  }
}
