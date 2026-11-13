/**
 * The bytes a PDF export renders.
 *
 * Contract: the renderer receives either the original bytes, or a temporary copy in which only the order of
 * child elements inside schema-ordered containers differs from the original. Every other part is
 * content-identical. If that cannot be established, the export is refused and nothing is rendered. The source
 * file is never written.
 *
 * Order of operations:
 *   0. Package limits (`checkZipBombCeilings`, headers only) for every input, before validation and rendering.
 *   1. Validate the original. Valid: the original bytes are used, with no XML parse and no copy.
 *   2. Invalid: repair a temporary copy by reordering only (no renames, no infills, no `DocxParser`, so no
 *      paragraph id is assigned and no other serialization happens). Only parts established as UTF-8 are
 *      examined; every other part is copied byte for byte.
 *   3. Prove each repaired part with `verifyOrderOnlyRepair` (an independent comparison).
 *   4. Build the copy, check every untouched part is content-identical, and validate the copy.
 *   5. Any doubt (a part that does not parse, a failed proof, a copy that still has errors) refuses with the
 *      validation errors and a remedy that says what to do.
 */
import { createHash } from 'crypto';
import AdmZip from '../adm-zip-safe';
import { OoxmlValidationError, validateDocxBuffer } from '../ooxml/validator-gate';
import { reorderPartXml, verifyOrderOnlyRepair } from '../ooxml/normalize';
import { assertNoXxeAdjacent } from '../ooxml/xxe-guard';
import { checkZipBombCeilings } from '../zip-hardening';
import { isRogueZipEntry } from '../sanitizer';

export interface PdfSourcePreparation {
  /** The bytes to render: the original buffer itself when it was valid, else the verified copy. */
  buffer: Buffer;
  /** True when `buffer` is a repaired copy. */
  repaired: boolean;
  /** Names of the parts whose child order was repaired. */
  repairedParts: string[];
  /** How many containers had their children reordered. */
  reorderedContainers: number;
}

export const PDF_SOURCE_REMEDY =
  'the file has schema errors that ASTra does not repair for export. Run "astra validate <file>" for the list; ' +
  'if "astra normalize" reports repairs, run it and export the result.';

const sha256 = (data: Buffer): string => createHash('sha256').update(data).digest('hex');

function isXmlPart(name: string): boolean {
  const lower = name.toLowerCase();
  return lower.endsWith('.xml') || lower.endsWith('.rels');
}

const STRICT_UTF8 = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true });

/** UTF-16 by byte order mark or by the first characters of an XML declaration as 16-bit units. */
function looksUtf16(data: Buffer): boolean {
  const n = data.length;
  if (n < 2) return false;
  if ((data[0] === 0xff && data[1] === 0xfe) || (data[0] === 0xfe && data[1] === 0xff)) return true;
  if (n >= 4 && data[0] === 0x3c && data[1] === 0x00 && data[2] === 0x3f && data[3] === 0x00) return true;
  if (n >= 4 && data[0] === 0x00 && data[1] === 0x3c && data[2] === 0x00 && data[3] === 0x3f) return true;
  return false;
}

/**
 * The text of a part that is established as UTF-8, or null. Established means: not UTF-16, the bytes decode
 * strictly (an invalid sequence is not turned into U+FFFD), and the XML declaration names UTF-8 or no encoding.
 * A UTF-8 byte order mark is kept as a character so that it is written back. A part for which null is returned is
 * never rewritten.
 */
export function establishedUtf8Text(data: Buffer): string | null {
  if (looksUtf16(data)) return null;
  let text: string;
  try {
    text = STRICT_UTF8.decode(data);
  } catch {
    return null;
  }
  const declared = /^\uFEFF?\s*<\?xml\b[^>]*?\bencoding\s*=\s*["']([^"']+)["']/i.exec(text);
  if (declared && declared[1].toUpperCase() !== 'UTF-8') return null;
  return text;
}

export async function prepareDocxForPdfExport(
  original: Buffer,
  context: string = 'PDF export source'
): Promise<PdfSourcePreparation> {
  // 0. Package limits, for every input and before the validator or the renderer starts. Headers only: nothing
  // is inflated. An unreadable package is left to the validator, which reports it. A breach throws `E_LIMIT`.
  let source: InstanceType<typeof AdmZip> | null = null;
  let unreadable = '';
  try {
    source = new AdmZip(original);
    source.getEntries(); // adm-zip reads the directory lazily: a duplicate entry name throws here
  } catch (error) {
    unreadable = error instanceof Error ? error.message : String(error);
  }
  if (source) checkZipBombCeilings(source, { sourceCompressedBytes: original.length });

  // 1. A valid original goes to the renderer untouched.
  const first = await validateDocxBuffer(original);
  if (!source) {
    // The validator may accept a package that cannot be read as a zip here (for example a duplicate entry name,
    // which `adm-zip` refuses). Nothing is rendered from it either way.
    const errors = first.ok ? [{ description: `the package cannot be read as a zip (${unreadable})` }] : first.errors;
    throw new OoxmlValidationError(errors, context, PDF_SOURCE_REMEDY);
  }
  if (first.ok) return { buffer: original, repaired: false, repairedParts: [], reorderedContainers: 0 };
  const refuse = (errors: typeof first.errors, cause?: unknown): never => {
    const error = new OoxmlValidationError(errors, context, PDF_SOURCE_REMEDY);
    if (cause !== undefined) (error as { cause?: unknown }).cause = cause;
    throw error;
  };

  let copy: Buffer;
  const repairedParts: string[] = [];
  let reorderedContainers = 0;
  try {
    // Entry hygiene: a duplicate entry name or an unsafe name is never repaired around.
    const names = new Set<string>();
    for (const entry of source.getEntries()) {
      if (names.has(entry.entryName) || isRogueZipEntry(entry.entryName)) return refuse(first.errors);
      names.add(entry.entryName);
    }

    // 2 and 3. Reorder-only repair of each established UTF-8 XML part, each proved against its original.
    const replacements = new Map<string, Buffer>();
    for (const entry of source.getEntries()) {
      const name = entry.entryName;
      if (entry.isDirectory || !isXmlPart(name)) continue;
      const xml = establishedUtf8Text(entry.getData());
      if (xml === null) continue; // not established as UTF-8: not examined, not changed
      assertNoXxeAdjacent(xml, name);
      const result = reorderPartXml(name, xml);
      if (result.reordered.length === 0) continue;
      if (!verifyOrderOnlyRepair(name, xml, result.xml)) {
        // The repair changed more than order, or the proof could not be made: refuse.
        return refuse(first.errors);
      }
      replacements.set(name, Buffer.from(result.xml, 'utf8'));
      repairedParts.push(name);
      reorderedContainers += result.containers.length;
    }
    if (replacements.size === 0) return refuse(first.errors);

    // 4. Build the copy and check every part that was not repaired is content-identical.
    const target = new AdmZip(original);
    for (const [name, data] of replacements) target.updateFile(name, data);
    copy = target.toBuffer();
    const check = new AdmZip(copy);
    const before = new Map(source.getEntries().filter((e) => !e.isDirectory).map((e) => [e.entryName, sha256(e.getData())]));
    const after = new Map(check.getEntries().filter((e) => !e.isDirectory).map((e) => [e.entryName, sha256(e.getData())]));
    if (before.size !== after.size) return refuse(first.errors);
    for (const [name, hash] of before) {
      if (replacements.has(name)) continue;
      if (after.get(name) !== hash) return refuse(first.errors);
    }
  } catch (error) {
    if (error instanceof OoxmlValidationError) throw error;
    return refuse(first.errors, error);
  }

  // 5. The copy must be valid; a mixed file (order errors plus real ones) is refused with what remains.
  const second = await validateDocxBuffer(copy);
  if (!second.ok) return refuse(second.errors);
  return { buffer: copy, repaired: true, repairedParts, reorderedContainers };
}
