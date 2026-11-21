import AdmZip from '../adm-zip-safe';
import { DOMParser } from '../ooxml/dom-parser';
import { findRegion, findSlot, hashNodes } from '../compose/template-canonical';
import { buildTemplate } from './index';
import type { PageSetup } from './lib';

/** The hashes of the placeholder body and the title slot that a template has when it is new on a page. */
export interface PristineHashes {
  region: string;
  slots: Record<string, string>;
}

const cache = new Map<string, PristineHashes | null>();

/**
 * Builds a template for a page and measures it. The compose shortcut compares a document's body with this, so a document is recognised as
 * untouched whatever its paper and margins are: the body of a template follows its page, and the hash follows with it. Null for a template
 * with no placeholder body (blank). The result is a pure function of the arguments, so it is cached.
 */
export function pristineHashes(template: string, page: PageSetup): PristineHashes | null {
  const key = `${template}|${page.width}|${page.height}|${page.top}|${page.bottom}|${page.left}|${page.right}|${page.header}|${page.footer}|${page.gutter ?? 0}|${page.gutterAtTop ? 1 : 0}`;
  const hit = cache.get(key);
  if (hit !== undefined) return hit;
  const xml = new AdmZip(buildTemplate(template, page)).readAsText('word/document.xml');
  const body = new DOMParser().parseFromString(xml, 'text/xml').getElementsByTagName('w:body')[0];
  const region = body ? findRegion(body) : null;
  let result: PristineHashes | null = null;
  if (region && region.nodes.length > 0) {
    const slots: Record<string, string> = {};
    const title = findSlot(body, 'title');
    if (title) slots.title = hashNodes(title.nodes);
    result = { region: hashNodes(region.nodes), slots };
  }
  cache.set(key, result);
  return result;
}
