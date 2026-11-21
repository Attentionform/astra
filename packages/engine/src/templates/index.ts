/**
 * The built-in templates, built when a document is made. Each builder takes the page (size and margins, in twips) and writes a complete,
 * deterministic .docx: the same inputs give the same bytes. A template is therefore available at any valid page size, and the compose
 * shortcut can rebuild the pristine body for the page that a document has.
 */
import { blank, invoice, letter, manuscript, memo, newsletter, proposal, researchPaper, technicalSpec } from './others';
import { report } from './report';
import type { PageSetup } from './lib';

import { DESIGN_VERSION } from './themes';
export { DESIGN_VERSION };
export { LETTER, A4, type PageSetup } from './lib';

type Builder = (page: PageSetup) => Buffer;

const BUILDERS: Record<string, Builder> = {
  blank,
  report,
  memo,
  letter,
  proposal,
  invoice,
  'technical-spec': technicalSpec,
  'research-paper': researchPaper,
  novel: manuscript,
  newsletter,
};

/** The names of the built-in templates. */
export const BUILDER_NAMES: readonly string[] = Object.keys(BUILDERS);

/** Builds a template for a page. Throws for a name that is not a template. */
export function buildTemplate(name: string, page: PageSetup): Buffer {
  const build = BUILDERS[name];
  if (!build) throw new Error(`unknown template '${name}'`);
  return build(page);
}

/** The design version of a template this engine builds, or undefined for a name that is not one. */
export function designVersionOf(name: string): number | undefined {
  return name in BUILDERS ? DESIGN_VERSION : undefined;
}
