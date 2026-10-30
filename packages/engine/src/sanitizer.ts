import { z } from 'zod';
import AdmZip from './adm-zip-safe';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { checkZipBombCeilings, clampZipEntryTimes } from './zip-hardening';
import { assertNoForbiddenMarkup, stripSvgActiveContent } from './ooxml/xxe-guard';

export { assertNoForbiddenMarkup } from './ooxml/xxe-guard';

/**
 * Sprint-3 proofs: export-time metadata sanitizer.
 *
 * Purges revision/author/EXIF/PII traces from a DOCX package on export:
 * - RSIDs (`w:rsid*` attrs in document/styles/numbering/comments/… plus the
 *   `<w:rsid w:val>` element list in word/settings.xml)
 * - Authors (`w:author` / `w:initials` on tracked changes & comments,
 *   `dc:creator` / `cp:lastModifiedBy` in core props, `Company`/`Manager` in
 *   app props, `w:author` in people/personList parts)
 * - EXIF (JPEG APP1 Exif/XMP + APP13 Photoshop headers; PNG textual
 *   author/software/description chunks + eXIf + XMP)
 * - PII (email / phone / SSN / credit-card patterns): core/app/custom string
 *   properties plus body prose in document/headers/footers/footnotes/endnotes
 *   and the comment family (comments/commentsExtended/people), with split-run
 *   joining inside paragraphs so PII straddling adjacent w:t runs is caught;
 *   element text in customXml items, charts, and style-family parts (field
 *   codes such as w:instrText included via the element-text pass); carrier
 *   attributes in the comment family (e.g. people.xml w15:userId) and in
 *   style-family name attributes (w:name w:val); drawing alternative text
 *   (descr/title/name on wp:docPr and pic:cNvPr companions) in document-family
 *   parts; attribute values in chart and customXml parts (attribute scope adds
 *   a labeled-only bare-digit phone rule for tel/telephone/phone/mobile/fax
 *   labels, so "Phone: 5551234567" redacts while unlabeled years/counts and
 *   body w:t prose keep the conservative formatted-only phone scope).
 *   M2 extension (labeled/contextual ONLY, never global): body w:t prose and
 *   element text additionally redact labeled bare-digit phones
 *   (tel/telephone/phone/mobile/fax/call/contact + 7-15 digits, e.g.
 *   "call 5551234567") and labeled contiguous-9 SSN ("SSN 123456789").
 *   RESIDUAL (honest, by prior decision): UNLABELED bare-digit runs and
 *   UNLABELED contiguous-9 runs are NEVER redacted — global bare-digit
 *   matching is rejected (false positives on IDs/counts/years). "5551234567"
 *   alone and "123456789" alone survive; only the labeled shapes above
 *   redact.
 *
 * Zero-retention flag: `zeroRetention: true` (or env ASTRA_ZERO_RETENTION=1)
 * guarantees the report never echoes original author/PII values (counts only)
 * and no original bytes are retained in memory beyond the purge pass.
 *
 * Opt-in closure flags (both default false, compose with zeroRetention):
 * - `acceptAllRevisions: true` accepts tracked changes (Word "accept all"
 *   semantics): `w:del`/`w:moveFrom` subtrees removed entirely (deleted /
 *   moved-away text gone, incl. nested `w:ins`/`w:moveTo`); `w:ins`/`w:moveTo`
 *   unwrapped (inner runs kept, wrapper dropped);
 *   `w:moveFromRangeStart/End` + `w:moveToRangeStart/End` (+ customXml
 *   move/ins/del range markers) dropped; `w:cellIns`/`w:cellMerge` markers
 *   dropped (cell kept) while `w:cellDel` drops its containing `w:tc`
 *   (deleted cell gone); `*PrChange` family (`rPrChange/pPrChange/`
 *   `sectPrChange/numberingChange` + `tblPrChange/trPrChange/tcPrChange/`
 *   `tPrChange`, plus any other `W:*Change`) dropped keeping current
 *   formatting. Applies to document/headers/footers and every other word
 *   part via the shared document-XML pass. Malformed XML under this flag
 *   throws `E_VALIDATION` (fail-closed, never a silent clean report).
 * - `stripComments: true` removes `word/comments.xml` (+ `word/commentsExtended.xml`
 *   and `word/people.xml` only when `word/comments.xml` is present — standalone
 *   `people.xml` from external Word co-authoring without comments is preserved
 *   to avoid over-deletion; kept adjuncts still get author/RSID purging),
 *   plus orphan `word/_rels/<part>.rels` for each dropped part, purges all
 *   `w:commentRangeStart` / `w:commentRangeEnd` / `w:commentReference`
 *   elements from document/headers/footers, and drops the matching `Override`
 *   entries in `[Content_Types].xml` plus the comment `Relationship` entries
 *   in `word/_rels/document.xml.rels` so no dangling references survive
 *   (schema validation stays green). Malformed XML under this flag throws
 *   `E_VALIDATION` (fail-closed).
 * Deliberately preserved (benign, no secret): `w14:paraId` / `w15:paraId`
 * paragraph IDs (8-hex ST_LongHexNumber correlation IDs, no author/PII
 * payload) survive both flags.
 *
 * Macro/active-content gate (refuse-default): `vbaProject.bin` by name,
 * `word/activeX/*` by directory, `oleObject*.bin`/`activeX*.bin` bins
 * corroborated by a package relationship, plus `.rels` files containing a
 * macro relationship (`vbaProject`/`oleObject`/`activeX`/control rels) are
 * detected via `findMacroPartNames`. Default (neither `allowMacros` nor
 * `stripMacros`) throws `E_VALIDATION` naming the parts with a preserve-vs-
 * strip remedy and changes zero bytes; `allowMacros` preserves them
 * byte-identical; `stripMacros` removes bins + companions + macro-only rel
 * files and sweeps macro rel elements / content-type overrides (removals in
 * `strippedMacros`, detections in `macroParts`). Both flags together throw.
 */

export const SANITIZER_DEFAULT_AUTHOR = '';

export const SanitizeOptionsSchema = z
  .object({
    zeroRetention: z.boolean().optional(),
    purgeRSIDs: z.boolean().optional(),
    purgeAuthors: z.boolean().optional(),
    purgeExif: z.boolean().optional(),
    purgePII: z.boolean().optional(),
    acceptAllRevisions: z.boolean().optional(),
    stripComments: z.boolean().optional(),
    replacementAuthor: z.string().max(256).optional(),
    allowMacros: z.boolean().optional(),
    stripMacros: z.boolean().optional(),
  })
  .strict();

export type SanitizeOptions = z.infer<typeof SanitizeOptionsSchema>;

export interface SanitizeReport {
  ok: true;
  zeroRetention: boolean;
  removedRSIDs: number;
  purgedAuthors: number;
  strippedExifBytes: number;
  purgedPII: number;
  bodyPII: number;
  propsPII: number;
  /** Revision nodes accepted (del/moveFrom/cellDel dropped, ins/moveTo unwrapped, range + cell + *Change markers resolved). */
  acceptedRevisions: number;
  /**
   * Comment parts (+ orphan part rels) removed + comment reference elements
   * purged + comment relationship / content-type entries removed
   * (stripComments). Counts only.
   */
  strippedComments: number;
  filesTouched: string[];
  /**
   * Content-scanned parts (sorted entry names): every package entry that
   * passed through a content-sanitizing scan in this run — document XML
   * parts (RSID/author/revision/comment/body-PII, including embedded
   * `word/embeddings/*.xml` and inner workbook sheet text), docProps core/app/custom
   * (author/PII), word/media/* binaries (EXIF gate), and
   * [Content_Types].xml / *.rels entries when the comment-trace strip ran.
   * Pass-through blobs (theme/fontTable, unexamined rels) and name-gate
   * drops (droppedRogue, dropped comment parts) are NOT listed here; they
   * have their own evidence lists. Purpose: a zero purge count alongside a
   * non-empty scannedParts reads as "scanned, none found" instead of
   * "maybe never scanned". Additive only — old JSON readers ignore it.
   */
  scannedParts: string[];
  /** Scalar mirror of scannedParts.length ("scanned:N" at a glance). */
  scannedCount: number;
  /**
   * Names of zip entries removed before sanitizing: traversal / absolute /
   * drive-letter / control-char rogue entries plus denied executable and
   * active entries (exe dll bat cmd ps1 psm1 vbs js jse ws wsf scr msi lnk
   * jar hta cpl com pif reg url,
   * case-insensitive on the final segment). Length is the drop count.
   * Choice (documented): silent drop+count, not a loud throw, so one
   * hostile entry cannot break export of an otherwise legit package and
   * opaque third-party parts keep flowing; visibility comes from this list
   * (names, never echoed bytes).
   */
  droppedRogue: string[];
  /**
   * Macro/active-content parts detected in the input package (sorted entry
   * names): `vbaProject.bin` by name, `word/activeX/*` cluster parts by
   * directory, `oleObject*.bin`/`activeX*.bin` bins corroborated by a package
   * relationship, plus `.rels` files containing a macro relationship.
   * Present in every report ([] when none) so a refuse/allow/strip decision
   * stays auditable. Additive only — old JSON readers ignore it.
   */
  macroParts: string[];
  /**
   * SVG media parts (`word/media/*.svg`) where script-capable constructs
   * (script, foreignObject, event-handler attributes, XInclude/XSLT/stylesheet
   * PI, javascript:/vbscript:/data:text/html hrefs, plus strip-only <style>)
   * were stripped by the sanitize path (D5 strip-loud contract: kinds are
   * reported by the xxe-guard helper, this list names the parts — never the
   * payload). Sorted entry names; [] when no SVG needed stripping. Clean SVGs
   * (plain shapes, https hrefs, data:image/* raster embeds, data:font/*)
   * pass through byte-identical. Compose ingest still refuses hostile SVGs
   * (assert path, untouched by this strip) while allowing <style> +
   * data:image/font. Additive only.
   */
  strippedSvgActive: string[];
  /**
   * Macro entries removed by `stripMacros` (sorted): macro bins, same-stem
   * companion XMLs, and macro-only rel files. Kept `.rels` files that only
   * lost Relationship elements (plus `[Content_Types].xml` override sweeps)
   * are reported via `filesTouched`, not here. [] unless stripping removed
   * entries. Additive only.
   */
  strippedMacros: string[];
}

export function validateSanitizeOptionsStrict(op: unknown): {
  ok: true;
  options: Required<SanitizeOptions>;
} {
  const parsed = SanitizeOptionsSchema.safeParse(op ?? {});
  if (!parsed.success) {
    const detail = parsed.error.issues
      .map((i) => `${i.path.join('.') || '(root)'}: ${i.message} [${i.code}]`)
      .join('; ');
    throw new Error(
      `E_VALIDATION: invalid sanitizer options. ${detail}. ` +
        `Remedy: pass a strict subset of { zeroRetention, purgeRSIDs, purgeAuthors, purgeExif, purgePII, acceptAllRevisions, stripComments, replacementAuthor, allowMacros, stripMacros }.`
    );
  }
  const d = parsed.data;
  if (d.allowMacros === true && d.stripMacros === true) throw macroChoiceError();
  return {
    ok: true,
    options: {
      zeroRetention: d.zeroRetention ?? isZeroRetentionEnabled(),
      purgeRSIDs: d.purgeRSIDs ?? true,
      purgeAuthors: d.purgeAuthors ?? true,
      purgeExif: d.purgeExif ?? true,
      purgePII: d.purgePII ?? true,
      acceptAllRevisions: d.acceptAllRevisions ?? false,
      stripComments: d.stripComments ?? false,
      replacementAuthor: d.replacementAuthor ?? SANITIZER_DEFAULT_AUTHOR,
      allowMacros: d.allowMacros ?? false,
      stripMacros: d.stripMacros ?? false,
    },
  };
}

/** Zero-retention mode: explicit option wins, otherwise env ASTRA_ZERO_RETENTION. */
export function isZeroRetentionEnabled(opts?: Pick<SanitizeOptions, 'zeroRetention'>): boolean {
  if (opts?.zeroRetention !== undefined) return opts.zeroRetention === true;
  const raw = (process.env.ASTRA_ZERO_RETENTION || '').trim().toLowerCase();
  return raw === '1' || raw === 'true' || raw === 'yes';
}

function resolveOptions(opts?: SanitizeOptions): Required<SanitizeOptions> {
  return validateSanitizeOptionsStrict(opts ?? {}).options;
}

// Matches any revision-save id attribute: w:rsidR, w:rsidRPr, w:rsidP,
// w:rsidRDefault, w:rsidTr, w:rsidTc, w:rsidSect, … (case-sensitive per OOXML).
const RSID_ATTR_RE = /\s+w:rsid[A-Za-z]*="[^"]*"/g;
// word/settings.xml carries the rsid history as elements, not attributes:
// <w:rsids><w:rsid w:val="00A1B2C3"/>…</w:rsids>. These carry no content value
// (document bytes are identical without them) and must be purged alongside
// the rsid* attributes above. Includes <w:rsidRoot w:val="…"/> (creation RSID).
const RSID_ELEM_RE = /<w:rsid(?:Root)?\s+w:val="[^"]*"\s*\/>/g;
// After purging rsid elements the <w:rsids> wrapper is left empty — strip it
// cleanly so no empty revision container leaks structure.
const RSIDS_EMPTY_RE = /<w:rsids>\s*<\/w:rsids>/g;
const RSIDS_SELF_CLOSING_RE = /<w:rsids\s*\/>/g;
// Tracked-change / comment authorship. Covers w:author plus w15:author used by
// Word 2013+/365 in word/people.xml (<w15:person w15:author="…">).
// F3: capture the prefix so the replacement preserves w: vs w15: (hardcoding
// w:author leaves an undeclared w15: prefix use / drops the w15 binding).
const AUTHOR_ATTR_RE = /\s+(w|w15):author="[^"]*"/g;
const INITIALS_ATTR_RE = /\s+w:initials="[^"]*"/g;

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
// Phone matcher (conservative scope, documented): covers leading-+ international,
// paren-grouped US, and bare US 555-123-4567 (dash/dot/space separated). It
// deliberately does NOT match plain 10-digit runs (e.g. 5551234567) or short
// fragments, to avoid mangling ordinary counts/IDs/years in document text.
// Statutory clause ids (DFARS/FAR/CFR/NIST 252.204-7012, including "clause"
// and section-sign forms) are NOT excluded here; they are filtered by
// isClauseCitedNumber at every phone site below. The filter fires only for
// clause-shaped runs (dot-then-dash, e.g. 252.204-7012), so real agency
// numbers after the same acronyms (NIST: 301-975-2000, DFARS: 703-697-7241)
// still redact. The leading-+ international and paren-grouped branches are
// never clause-shaped (clause ids never render parenthesized or with a `+`
// prefix), so they always redact — avoiding leaks near acronyms.
const PHONE_RE = /(\+\d[\d\s\-().]{6,}\d|\(\d{3}\)\s*\d{3}[\s\-.]\d{4}|\b\d{3}[\s\-.]\d{3}[\s\-.]\d{4}\b)/g;

// Clause-citation guard: acronyms whose nearby numbers may be statutory
// clause ids rather than phone numbers.
const CLAUSE_ACRONYM_RE = /\b(dfars|far|cfr|nist)\b/gi;
// Allowed gap between the acronym and the number: separators plus clause
// words, prior clause numbers (for "252.204-7012, 252.204-7013" lists),
// and list joints. Anything else (notably dash-only digit runs) breaks the
// guard so real phone numbers keep redacting.
const CLAUSE_GAP_RE =
  /^(?:[\s.:;§\-–—,()]*|\b(?:clause|clauses|section|sec\.?|no\.?|number|and)\b[\s.:;§\-–—,()]*|\b\d{2,3}\.\d{2,4}-\d{2,4}\b[\s.:;§\-–—,()]*)*$/i;
// Clause ids join groups with a dot then a dash (252.204-7012, 52.204-21);
// phone-shaped runs (dashes/dots throughout, no dot-dash pair) never match.
const CLAUSE_NUMBER_RE = /\b\d{2,3}\.\d{2,4}-\d{2,4}\b/;

/** True when the digit run at [start, end) is a statutory clause citation. */
function isClauseCitedNumber(text: string, start: number, end: number): boolean {
  if (!CLAUSE_NUMBER_RE.test(text.slice(start, end))) return false;
  const before = text.slice(Math.max(0, start - 64), start);
  CLAUSE_ACRONYM_RE.lastIndex = 0;
  let lastEnd = -1;
  let m: RegExpExecArray | null;
  while ((m = CLAUSE_ACRONYM_RE.exec(before)) !== null) {
    lastEnd = m.index + m[0].length;
  }
  CLAUSE_ACRONYM_RE.lastIndex = 0;
  if (lastEnd < 0) return false;
  return CLAUSE_GAP_RE.test(before.slice(lastEnd));
}

/** Phone redaction with the clause-citation guard applied. Counts only redacted runs. */
function redactPhonesWithClauseGuard(text: string, redact: boolean): { text: string; count: number } {
  const rx = new RegExp(PHONE_RE.source, 'g');
  const hits: Array<{ s: number; e: number }> = [];
  let m: RegExpExecArray | null;
  while ((m = rx.exec(text)) !== null) {
    if (m[0].length === 0) {
      rx.lastIndex++;
      continue;
    }
    const s = m.index;
    const e = m.index + m[0].length;
    if (!isClauseCitedNumber(text, s, e)) hits.push({ s, e });
  }
  if (!redact) return { text, count: hits.length };
  if (hits.length === 0) return { text, count: 0 };
  let out = '';
  let pos = 0;
  for (const h of hits) {
    out += text.slice(pos, h.s) + '[redacted]';
    pos = h.e;
  }
  out += text.slice(pos);
  return { text: out, count: hits.length };
}

// Labeled-only bare-digit phones (attribute carriers + body prose, never
// global): contiguous 7-15 digit runs redacted only when directly introduced
// by a tel/telephone/phone/mobile/fax/call/contact label in the same value
// (e.g. "Phone: 5551234567", "fax 5551234", "call 5551234567"). Bare-digit
// runs without such a label are never matched here, so years (2026), counts,
// and structural ids stay intact. Disjoint from PHONE_RE (which requires
// separators, parens, or a leading +), so no double-count when both passes
// run. Global bare-digit matching is REJECTED by prior decision (false
// positives on IDs/counts) — this labeled/contextual rule is the entire
// bare-digit scope; unlabeled digit runs are a documented residual.
const LABELED_PHONE_RE = /(\b(?:tel|telephone|phone|mobile|fax)\b\s*[:：]?\s*\+?)(\d{7,15})\b/gi;

// Body-prose labeled bare-digit phones: same rule as attributes plus
// call/contact conversational labels ("call 5551234567"). Years (2026) and
// short counts never match (7-15 digits required + label prefix).
const LABELED_BARE_PHONE_BODY_RE =
  /(\b(?:tel|telephone|phone|mobile|fax|call|contact)\b\s*[:#：]?\s*\+?)(\d{7,15})\b/gi;

// Body-prose labeled contiguous-9 SSN: "SSN 123456789" / "Social Security
// 123456789". Dashed 123-45-6789 is already covered globally by SSN_RE, so
// this matches contiguous-9 ONLY (no double-count on dashed). Unlabeled
// 9-digit runs are a documented residual (global matching rejected: IDs).
const LABELED_SSN_BODY_RE =
  /(\b(?:ssn|social\s*security)(?:\s*(?:number|no\.?))?\s*[:#：]?\s*)(\d{9})\b/gi;

/** Labeled bare-digit redaction (attribute values only). Keeps the label, redacts the digits. */
function redactLabeledPhones(text: string, redact: boolean): { text: string; count: number } {
  const rx = new RegExp(LABELED_PHONE_RE.source, 'gi');
  let count = 0;
  const out = text.replace(rx, (_m: string, prefix: string, digits: string) => {
    void digits;
    count++;
    return redact ? `${prefix}[redacted]` : _m;
  });
  // Counting pass must report hits even when redact is false (replace ran
  // with identity above, so count is exact either way).
  if (!redact) return { text, count };
  return { text: out, count };
}

/**
 * M2 body-prose labeled bare-digit redaction (w:t prose + element text,
 * never global): labeled bare phones ("call 5551234567") + labeled
 * contiguous-9 SSN ("SSN 123456789"). Keeps the label, redacts the digits.
 * Unlabeled years/counts/IDs never match (label prefix required).
 */
function redactLabeledBareBody(text: string, redact: boolean): { text: string; count: number } {
  let count = 0;
  let out = text;
  const phoneRx = new RegExp(LABELED_BARE_PHONE_BODY_RE.source, 'gi');
  out = out.replace(phoneRx, (_m: string, prefix: string, digits: string) => {
    void digits;
    count++;
    return redact ? `${prefix}[redacted]` : _m;
  });
  const ssnRx = new RegExp(LABELED_SSN_BODY_RE.source, 'gi');
  out = out.replace(ssnRx, (_m: string, prefix: string, digits: string) => {
    void digits;
    count++;
    return redact ? `${prefix}[redacted]` : _m;
  });
  if (!redact) return { text, count };
  return { text: out, count };
}

/** Digit-only spans of labeled bare-digit hits in `joined` (prefix excluded, digits only). */
function labeledBareDigitSpans(joined: string): Array<{ s: number; e: number }> {
  const spans: Array<{ s: number; e: number }> = [];
  for (const src of [LABELED_BARE_PHONE_BODY_RE.source, LABELED_SSN_BODY_RE.source]) {
    const rx = new RegExp(src, 'gi');
    let m: RegExpExecArray | null;
    while ((m = rx.exec(joined)) !== null) {
      const prefix = m[1] ?? '';
      const digits = m[2] ?? '';
      if (!digits) continue;
      const ds = m.index + prefix.length;
      spans.push({ s: ds, e: ds + digits.length });
      if (m[0].length === 0) rx.lastIndex++;
    }
  }
  return spans.sort((a, b) => a.s - b.s);
}
// SSN (US): 123-45-6789. Distinct length (9 digits, 3-2-4) — no overlap with
// PHONE (10 digits) or CREDIT_CARD (13-19 digits).
const SSN_RE = /\b\d{3}-\d{2}-\d{4}\b/g;
// Credit card: 13-19 digits, contiguous or spaced/dashed grouped
// (e.g. 4111111111111111, 4111 1111 1111 1111, 4111-1111-1111-1111).
// Starts/ends on a digit; single space/dash allowed between digits.
// Deliberately redacts any 13-19 digit run (Luhn-lite alternative: over-redact
// rather than leak PII on non-Luhn test numbers).
const CREDIT_CARD_RE = /\b\d(?:[ -]?\d){12,18}\b/g;

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Central XXE / DTD detector lives in ooxml/xxe-guard (single pattern for all ingress points). */

/** P0-6: Reject / strip rogue zip entries containing parent traversal segments, absolute paths, drive letters, or control chars */
export function isRogueZipEntry(entryName: string): boolean {
  const trimmed = entryName.trim();
  const normalized = trimmed.replace(/\\/g, '/');
  // '..' blocks only as a full path segment (a/../b, ../x) so legit dotted
  // filenames (my..photo.png, report...final.docx) survive while traversal dies.
  const hasParentSegment = normalized.split('/').some((seg) => seg === '..');
  if (
    trimmed.startsWith('/') ||
    trimmed.startsWith('\\') ||
    /^[a-zA-Z]:/.test(trimmed) ||
    hasParentSegment ||
    // eslint-disable-next-line no-control-regex -- intentional: rogue-entry gate must detect C0 controls + DEL in zip entry names
    /[\x00-\x1f\x7f]/.test(entryName)
  ) {
    return true;
  }
  return false;
}

function isXmlPart(name: string): boolean {
  const lower = name.toLowerCase();
  return lower.endsWith('.xml') || lower.endsWith('.rels') || lower === '[content_types].xml';
}

/**
 * Executable / active entries never belong in a docx payload (macros live in
 * vbaProject.bin, not in loose exe/js parts). Deny-list on the final segment
 * after the last / (backslash-tolerant), case-insensitive, so
 * evil/malware.exe, evil/TOOL.BAT, and nested drops all match while
 * word/media/image.exe.png (final ext png) passes. Choice (documented):
 * same silent drop+count path as rogue entries, not a loud throw — a loud
 * throw would break export of legit packages that merely carry an opaque
 * third-party part, and the report list keeps the drop visible. The list is
 * intentionally narrow (known execution vectors); completeness is a known
 * gap, see report.
 */
const BLOCKED_EXECUTABLE_EXTS = new Set([
  'exe',
  'dll',
  'bat',
  'cmd',
  'ps1',
  'psm1',
  'vbs',
  'js',
  'jse',
  'ws',
  'wsf',
  'scr',
  'msi',
  'lnk',
  'jar',
  'hta',
  'cpl',
  'com',
  'pif',
  'reg',
  'url',
]);

export function isBlockedExecutableEntry(entryName: string): boolean {
  const afterSlash = entryName.split('/').pop() ?? '';
  const segment = afterSlash.split('\\').pop() ?? '';
  const dot = segment.lastIndexOf('.');
  if (dot < 0) return false;
  return BLOCKED_EXECUTABLE_EXTS.has(segment.slice(dot + 1).toLowerCase());
}

/** Single drop predicate: rogue path shapes plus denied executables. */
function isDroppedEntry(entryName: string): boolean {
  return isRogueZipEntry(entryName) || isBlockedExecutableEntry(entryName);
}

/**
 * Macro/active-content gate (refuse-default).
 *
 * Detection rule (documented, conservative so legit embeddings survive):
 * - `vbaProject.bin` (exact basename, case-insensitive) is macro by NAME
 *   alone — a VBA project part is never legit document content.
 * - Everything under `word/activeX/` (bins, control XMLs, their rels) is
 *   macro by DIRECTORY alone — that cluster only exists for ActiveX controls.
 * - `oleObject*.bin` / `activeX*.bin` bins outside `word/embeddings/` and
 *   outside the shapes below count only when CORROBORATED by a package
 *   relationship targeting them (an unreferenced `word/oleObject1.bin` stays
 *   a legit opaque part, matching the long-standing R-10 passthrough).
 *   EXCEPTION (M4): `word/embeddings/*.xlsm|*.xlsb|*.bin` are macro by NAME
 *   alone (no corroboration) — opaque binaries the content scans cannot see
 *   (only *.xlsx + *.xml are scanned). Same refuse-default disposition as
 *   vbaProject. Legit *.xlsx workbooks are unaffected.
 * - A `.rels` file counts when it contains a macro relationship: a
 *   `Type` naming `vbaProject` / `oleObject` / `activeX` (or the ActiveX
 *   control rel type), or a `Target` naming a macro bin above.
 *
 * Disposition: default (neither flag) REFUSES with `E_VALIDATION` naming the
 * parts; `allowMacros` preserves bytes verbatim; `stripMacros` removes macro
 * bins + same-stem companion XMLs + macro-only rel files and sweeps macro
 * Relationship elements / content-type overrides from kept files. Both flags
 * together are a usage error (refuse, choose one).
 */

/** Final path segment of a zip entry name (backslash-tolerant). */
function macroBasename(entryName: string): string {
  const afterSlash = entryName.split('/').pop() ?? '';
  return afterSlash.split('\\').pop() ?? '';
}

/** Exact `vbaProject.bin` basename (case-insensitive). */
function isVbaProjectBinName(basename: string): boolean {
  return basename.toLowerCase() === 'vbaproject.bin';
}

/** Whole `word/activeX/` cluster (bins, control XMLs, their rels). */
function isActiveXPartName(entryName: string): boolean {
  return /^word\/activex\//i.test(entryName);
}

/** `oleObject*.bin` / `activeX*.bin` basename shape (needs rel corroboration). */
function isCorroboratedMacroBinName(basename: string): boolean {
  const lower = basename.toLowerCase();
  return lower.endsWith('.bin') && (lower.startsWith('oleobject') || lower.startsWith('activex'));
}

/**
 * M4 opaque embedding binaries: `word/embeddings/*.xlsm|*.xlsb|*.bin` are
 * macro by NAME alone (no relationship corroboration required). Rationale:
 * these are opaque binary blobs the sanitizer cannot scan (only *.xlsx +
 * *.xml get content scans) — macro-capable spreadsheet binaries (xlsm/xlsb)
 * and unidentified .bin payloads must not pass opaque by default. Same
 * refuse-default disposition as vbaProject (E_VALIDATION naming the parts;
 * --allow-macros preserves byte-identical, --strip-macros removes). Legit
 * `*.xlsx` workbooks are NOT in this set and keep flowing (scanned, intact).
 */
function isOpaqueEmbeddingMacroName(entryName: string): boolean {
  const lower = entryName.toLowerCase();
  if (!lower.startsWith('word/embeddings/')) return false;
  return lower.endsWith('.xlsm') || lower.endsWith('.xlsb') || lower.endsWith('.bin');
}

/** ActiveX control relationship type (points at the control XML companion). */
const ACTIVEX_CONTROL_REL_TYPE = 'http://schemas.microsoft.com/office/2006/relationships/control';

/** True for macro relationship Types (case-insensitive token match). */
function isMacroRelType(type: string): boolean {
  const lower = (type || '').toLowerCase();
  return (
    lower.includes('vbaproject') ||
    lower.includes('oleobject') ||
    lower.includes('activex') ||
    lower === ACTIVEX_CONTROL_REL_TYPE
  );
}

interface ParsedMacroRel {
  type: string;
  target: string;
  targetBase: string;
}

const MACRO_REL_TAG_RE = /<Relationship\b[^>]*>/g;

function parseMacroRels(text: string): ParsedMacroRel[] {
  const out: ParsedMacroRel[] = [];
  if (!text.includes('Relationship')) return out;
  const re = new RegExp(MACRO_REL_TAG_RE.source, 'g');
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const tag = m[0];
    const typeMatch = /Type="([^"]*)"/.exec(tag);
    const targetMatch = /Target="([^"]*)"/.exec(tag);
    const target = targetMatch ? targetMatch[1] : '';
    out.push({
      type: typeMatch ? typeMatch[1] : '',
      target,
      targetBase: macroBasename(target),
    });
  }
  return out;
}

interface MacroEvidence {
  /** Macro parts (bins + cluster files + rels files containing a macro rel). */
  parts: string[];
  /** Parsed relationships per rels file (only files that were readable). */
  relsByFile: Map<string, ParsedMacroRel[]>;
}

/** Collects macro evidence over entry names (regex-only, never throws). */
function collectMacroEvidence(
  entryNames: Iterable<string>,
  readText?: (name: string) => string | null
): MacroEvidence {
  const names = [...entryNames];
  const relsByFile = new Map<string, ParsedMacroRel[]>();
  if (readText) {
    for (const n of names) {
      if (!n.toLowerCase().endsWith('.rels')) continue;
      let text: string | null = null;
      try {
        text = readText(n);
      } catch {
        text = null;
      }
      if (!text) continue;
      relsByFile.set(n, parseMacroRels(text));
    }
  }
  const referencedBases = new Set<string>();
  for (const rels of relsByFile.values()) {
    for (const r of rels) {
      if (r.targetBase) referencedBases.add(r.targetBase);
    }
  }
  const macroBins = new Set<string>();
  for (const n of names) {
    const base = macroBasename(n);
    if (isVbaProjectBinName(base) || isActiveXPartName(n)) {
      macroBins.add(n);
    } else if (isOpaqueEmbeddingMacroName(n)) {
      macroBins.add(n);
    } else if (isCorroboratedMacroBinName(base) && referencedBases.has(base)) {
      macroBins.add(n);
    }
  }
  const macroBinBases = new Set<string>();
  for (const n of macroBins) macroBinBases.add(macroBasename(n));
  const parts = new Set<string>(macroBins);
  for (const [file, rels] of relsByFile) {
    for (const r of rels) {
      if ((r.type && isMacroRelType(r.type)) || (r.targetBase && macroBinBases.has(r.targetBase))) {
        parts.add(file);
        break;
      }
    }
  }
  return { parts: [...parts].sort(), relsByFile };
}

/**
 * Lists macro/active-content part names (sorted): macro bins + ActiveX
 * cluster files + `.rels` files containing a macro relationship. [] when
 * the package carries no active content. Never throws (unreadable rels
 * simply contribute no corroboration).
 */
export function findMacroPartNames(
  entryNames: Iterable<string>,
  readText?: (name: string) => string | null
): string[] {
  return collectMacroEvidence(entryNames, readText).parts;
}

/** Refuse-default error: names the parts, offers preserve vs strip. */
export function macroRefusalError(parts: readonly string[], verb = 'process'): Error {
  const names = [...parts].sort();
  const shown = names.slice(0, 10).join(', ') + (names.length > 10 ? `, +${names.length - 10} more` : '');
  const remedy =
    'Pass --allow-macros to preserve macro parts and continue, or --strip-macros to remove macro parts and continue (never both).';
  const err = new Error(
    `E_VALIDATION: refusing to ${verb} document with macro/active content parts (${shown}). Remedy: ${remedy}`
  );
  (err as unknown as { code?: string }).code = 'E_VALIDATION';
  (err as unknown as { remedy?: string }).remedy = remedy;
  return err;
}

/** Both-flags usage error (preserve vs strip is an exclusive choice). */
export function macroChoiceError(): Error {
  const remedy = 'Pass exactly one of --allow-macros (preserve macro parts) or --strip-macros (remove macro parts).';
  const err = new Error(`E_VALIDATION: --allow-macros and --strip-macros are mutually exclusive. Remedy: ${remedy}`);
  (err as unknown as { code?: string }).code = 'E_VALIDATION';
  (err as unknown as { remedy?: string }).remedy = remedy;
  return err;
}

/** Minimal entry-store seam so one strip routine serves Maps and live zips. */
export interface MacroEntryStore {
  names(): string[];
  readText(name: string): string | null;
  remove(name: string): void;
  writeText(name: string, text: string): void;
}

/** Adapts an in-memory package map to the strip seam. */
export function mapMacroStore(files: Map<string, Buffer>): MacroEntryStore {
  return {
    names: () => [...files.keys()],
    readText: (name: string) => {
      const buf = files.get(name);
      return buf ? buf.toString('utf8') : null;
    },
    remove: (name: string) => {
      files.delete(name);
    },
    writeText: (name: string, text: string) => {
      files.set(name, Buffer.from(text, 'utf8'));
    },
  };
}

/**
 * Removes macro Relationship elements targeting stripped parts (or carrying
 * a macro rel type) from a kept `.rels` part. Mirrors the comment-rel sweep;
 * malformed XML throws `E_VALIDATION` (fail-closed).
 */
function stripMacroRelationshipElements(
  xml: string,
  removedBasenames: Set<string>,
  context = 'rels part'
): { xml: string; removed: number } {
  assertNoForbiddenMarkup(xml, context);
  const none = { xml, removed: 0 };
  if (!xml.includes('Relationship')) return none;
  const { head, body } = splitXmlProlog(xml);
  const doc = parseFailClosedXml(body, context);
  const list = doc.getElementsByTagNameNS(REL_NS, 'Relationship');
  const doomed: Element[] = [];
  for (let i = 0; i < list.length; i++) {
    const el = list.item(i) as Element;
    const type = el.getAttribute('Type') ?? '';
    const target = el.getAttribute('Target') ?? '';
    const base = macroBasename(target);
    if ((type && isMacroRelType(type)) || (base && removedBasenames.has(base))) doomed.push(el);
  }
  if (doomed.length === 0) return none;
  for (const el of doomed) el.parentNode?.removeChild(el);
  return { xml: serializeWithProlog(head, doc), removed: doomed.length };
}

/**
 * Removes `[Content_Types].xml` Override elements for stripped macro parts
 * (matched by PartName, any content type). Mirrors the comment-CT sweep.
 */
function stripMacroContentTypeOverrides(
  xml: string,
  removedPartNames: Set<string>
): { xml: string; removed: number } {
  assertNoForbiddenMarkup(xml, '[Content_Types].xml');
  const none = { xml, removed: 0 };
  if (removedPartNames.size === 0 || !xml.includes('Override')) return none;
  const { head, body } = splitXmlProlog(xml);
  const doc = parseFailClosedXml(body, '[Content_Types].xml');
  const list = doc.getElementsByTagNameNS(CT_NS, 'Override');
  const doomed: Element[] = [];
  for (let i = 0; i < list.length; i++) {
    const el = list.item(i) as Element;
    const partName = el.getAttribute('PartName');
    if (partName && removedPartNames.has(partName)) doomed.push(el);
  }
  if (doomed.length === 0) return none;
  for (const el of doomed) el.parentNode?.removeChild(el);
  return { xml: serializeWithProlog(head, doc), removed: doomed.length };
}

/**
 * Strips macro parts through the entry-store seam: macro bins, same-stem
 * companion XMLs beside a removed bin, and macro-only rel files (every
 * Relationship targets a removed part) are deleted; kept `.rels` files lose
 * their macro Relationship elements and `[Content_Types].xml` loses the
 * matching Overrides so no dangling reference survives.
 */
export function applyMacroStrip(store: MacroEntryStore): { removed: string[]; edited: string[] } {
  const names = store.names();
  const nameSet = new Set(names);
  const { parts, relsByFile } = collectMacroEvidence(names, (n) => store.readText(n));
  const partSet = new Set(parts);
  const remove = new Set<string>();
  for (const n of names) {
    if (partSet.has(n) && !n.toLowerCase().endsWith('.rels')) remove.add(n);
  }
  for (const b of [...remove]) {
    if (!b.toLowerCase().endsWith('.bin')) continue;
    const slash = b.lastIndexOf('/');
    const dir = slash >= 0 ? b.slice(0, slash + 1) : '';
    const stem = (slash >= 0 ? b.slice(slash + 1) : b).replace(/\.bin$/i, '');
    const companion = `${dir}${stem}.xml`;
    if (nameSet.has(companion)) remove.add(companion);
  }
  const removeBases = new Set<string>();
  for (const n of remove) {
    const b = macroBasename(n);
    if (b) removeBases.add(b);
  }
  for (const [file, rels] of relsByFile) {
    if (remove.has(file) || rels.length === 0) continue;
    if (rels.every((r) => r.targetBase !== '' && removeBases.has(r.targetBase))) remove.add(file);
  }
  for (const n of remove) {
    try {
      store.remove(n);
    } catch {
      /* best-effort delete; a surviving entry is re-caught by the sweep below */
    }
  }
  const removed = [...remove].sort();
  const edited: string[] = [];
  for (const [file, rels] of relsByFile) {
    if (remove.has(file) || rels.length === 0) continue;
    const cur = store.readText(file);
    if (!cur) continue;
    const r = stripMacroRelationshipElements(cur, removeBases, file);
    if (r.xml !== cur) {
      store.writeText(file, r.xml);
      edited.push(file);
    }
  }
  if (nameSet.has('[Content_Types].xml') && !remove.has('[Content_Types].xml')) {
    const cur = store.readText('[Content_Types].xml');
    if (cur) {
      const withSlash = new Set(removed.map((n) => `/${n}`));
      const r = stripMacroContentTypeOverrides(cur, withSlash);
      if (r.xml !== cur) {
        store.writeText('[Content_Types].xml', r.xml);
        edited.push('[Content_Types].xml');
      }
    }
  }
  return { removed, edited: edited.sort() };
}

function isBodyPartName(name: string): boolean {
  return (
    name === 'word XML part' ||
    name === 'word/document.xml' ||
    /^word\/(header|footer|footnotes|endnotes).*\.xml$/.test(name)
  );
}

/** Comment-family parts: w:t paragraphs (split-run pass) + element text + carrier attributes. */
function isCommentFamilyPartName(name: string): boolean {
  return (
    name === 'word/comments.xml' ||
    name === 'word/commentsExtended.xml' ||
    name === 'word/people.xml'
  );
}

/** Embedded XML parts: ole companions and workbook XML snapshots beside bins. */
function isEmbeddingXmlPartName(name: string): boolean {
  return /^word\/embeddings\/.*\.xml$/i.test(name);
}

/** Embedded workbook packages: SpreadsheetML zips carrying sheet text. */
function isEmbeddingWorkbookPartName(name: string): boolean {
  return /^word\/embeddings\/.*\.xlsx$/i.test(name);
}

/** Non-w:t text carriers: customXml items, charts, embeddings, style-family parts. */
function isGenericTextPiiPartName(name: string): boolean {
  return (
    /^word\/customXml\/.*\.xml$/.test(name) ||
    /^word\/charts\/.*\.xml$/.test(name) ||
    isEmbeddingXmlPartName(name) ||
    isStyleFamilyPartName(name)
  );
}

/** Style-family parts: prose lives in name attributes (w:name/w:aliases w:val). */
function isStyleFamilyPartName(name: string): boolean {
  return (
    name === 'word/styles.xml' ||
    name === 'word/stylesWithEffects.xml' ||
    name === 'word/numbering.xml'
  );
}

/**
 * Redacts PII inside element text nodes (`>text<`) for parts that do not
 * carry prose in `w:t` (charts `a:t`/`c:v`, customXml items, style text).
 * Tag markup and attribute values are never touched here.
 * Runs of `w:t` prose are skipped here because they belong to the
 * split-run aware `w:t` pass; scanning them in both places counts one
 * redaction twice when only counting and would redact twice otherwise.
 */
export function redactTextNodePii(xml: string, redact: boolean): { xml: string; count: number } {
  if (!xml.includes('>') || !xml.includes('<')) return { xml, count: 0 };
  const WT_SKIP_RE = /<w:t\b[^>]*>[\s\S]*?<\/w:t>/g;
  const scanGap = (gap: string): { text: string; count: number } => {
    let gapCount = 0;
    const out = gap.replace(/>([^<>]+)</g, (_match, text: string) => {
      // M2: element text takes the body scanner (labeled bare digits
      // included); unlabeled years/counts stay intact.
      const { text: clean, count } = scanAndRedactBodyText(text, redact);
      gapCount += count;
      return `>${clean}<`;
    });
    return { text: out, count: gapCount };
  };
  let totalCount = 0;
  let out = '';
  let last = 0;
  let m: RegExpExecArray | null;
  const re = new RegExp(WT_SKIP_RE.source, 'g');
  while ((m = re.exec(xml)) !== null) {
    const gap = scanGap(xml.slice(last, m.index));
    totalCount += gap.count;
    out += gap.text + m[0];
    last = m.index + m[0].length;
  }
  const tail = scanGap(xml.slice(last));
  totalCount += tail.count;
  out += tail.text;
  return { xml: out, count: totalCount };
}

/**
 * Redacts PII inside double-quoted attribute values (e.g. people.xml
 * `w15:userId="jane@example.com"`). Element text and tag structure untouched.
 * Attribute scope adds the labeled-only bare-digit phone rule on top of the
 * shared email / formatted-phone / SSN / card patterns, so values such as
 * `descr="Phone: 5551234567"` redact while bare years/counts without a
 * label stay intact. Body `w:t` prose never takes this path.
 */
export function redactAttributePii(xml: string, redact: boolean): { xml: string; count: number } {
  if (!xml.includes('="')) return { xml, count: 0 };
  let totalCount = 0;
  const out = xml.replace(/="([^"]*)"/g, (_match, value: string) => {
    const { text: clean, count } = scanAndRedactAttributeValue(value, redact);
    totalCount += count;
    return `="${clean}"`;
  });
  return { xml: out, count: totalCount };
}

/**
 * Attribute-value scanner: shared patterns plus the labeled-only bare-digit
 * phone rule. Used by every attribute pass (comment/style carriers,
 * drawing alternative text, chart and customXml attributes). Body text
 * keeps the shared scanner without the labeled rule.
 */
function scanAndRedactAttributeValue(text: string, redact: boolean): { text: string; count: number } {
  const base = scanAndRedactPii(text, redact);
  const labeled = redactLabeledPhones(base.text, redact);
  return { text: labeled.text, count: base.count + labeled.count };
}

/** True for chart, customXml, and embedding XML parts (attribute values may carry prose). */
function isChartOrCustomXmlPartName(name: string): boolean {
  return (
    /^word\/customXml\/.*\.xml$/.test(name) ||
    /^word\/charts\/.*\.xml$/.test(name) ||
    isEmbeddingXmlPartName(name)
  );
}

/**
 * Redacts PII inside drawing alternative-text attributes (`descr`/`title`/
 * `name` on `wp:docPr` and `pic:cNvPr`/`cNvPr` companions) in document-family
 * parts. Only those three attributes on those tags are touched; all other
 * markup and structural attributes (ids, geometry) pass through untouched.
 * Field-code text (`w:instrText`) is already covered by the element-text
 * pass for the same parts and needs no separate step here.
 */
export function redactDrawingAltPii(xml: string, redact: boolean): { xml: string; count: number } {
  if (!xml.includes('docPr') && !xml.includes('cNvPr')) return { xml, count: 0 };
  if (!xml.includes('="')) return { xml, count: 0 };
  let totalCount = 0;
  const out = xml.replace(/<(?:wp:docPr|pic:cNvPr|cNvPr)\b[^>]*>/g, (tag: string) => {
    return tag.replace(/\s+(descr|title|name)="([^"]*)"/g, (_m: string, attr: string, value: string) => {
      const { text: clean, count } = scanAndRedactAttributeValue(value, redact);
      totalCount += count;
      return ` ${attr}="${clean}"`;
    });
  });
  return { xml: out, count: totalCount };
}

function scanAndRedactPii(text: string, redact: boolean): { text: string; count: number } {
  let count = 0;
  let out = text;

  const emailRe = new RegExp(EMAIL_RE.source, 'g');
  const emails = out.match(emailRe);
  if (emails) {
    count += emails.length;
    if (redact) {
      out = out.replace(emailRe, '[redacted]');
    }
  }

  const phoneRes = redactPhonesWithClauseGuard(out, redact);
  out = phoneRes.text;
  count += phoneRes.count;

  const ssnRe = new RegExp(SSN_RE.source, 'g');
  const ssns = out.match(ssnRe);
  if (ssns) {
    count += ssns.length;
    if (redact) {
      out = out.replace(ssnRe, '[redacted]');
    }
  }

  const cardRe = new RegExp(CREDIT_CARD_RE.source, 'g');
  const cards = out.match(cardRe);
  if (cards) {
    count += cards.length;
    if (redact) {
      out = out.replace(cardRe, '[redacted]');
    }
  }

  return { text: out, count };
}

/**
 * M2 body-text scanner: shared patterns plus labeled-only bare-digit rules
 * (labeled bare phones + labeled contiguous-9 SSN). Used by w:t fallback
 * runs outside any w:p and by non-w:t element text. Unlabeled years/counts
 * never match (label prefix required on the bare-digit branch).
 */
function scanAndRedactBodyText(text: string, redact: boolean): { text: string; count: number } {
  const base = scanAndRedactPii(text, redact);
  const labeled = redactLabeledBareBody(base.text, redact);
  return { text: labeled.text, count: base.count + labeled.count };
}

/** P0-7: Redact PII matches inside <w:t> tags in document body parts.
 *
 * Split-run evasion closure: Word freely splits one visible string across
 * adjacent runs (`<w:t>jane.doe@exa</w:t>…<w:t>mple.com</w:t>`), so per-run
 * scanning misses PII straddling a run boundary. This pass concatenates the
 * `w:t` texts within each `w:p`, scans the joined paragraph text, and maps
 * every match back onto the covered run slices (first slice keeps a single
 * `[redacted]` marker, continuation slices are cleared). Single-run matches
 * behave exactly as before. A fallback per-`w:t` pass covers runs outside any
 * `w:p` (no paragraph context to join). `redact: false` counts only.
 */
export function redactBodyPiiInWordText(xml: string, redact: boolean): { xml: string; count: number } {
  if (!xml.includes('<w:t')) return { xml, count: 0 };
  let totalCount = 0;
  const P_RE = /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/g;
  const WT_RE = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;
  const joined = xml.replace(P_RE, (_match, openTag: string, inner: string, closeTag: string) => {
    const { text, count } = redactSplitRunPiiInParagraph(inner, redact);
    totalCount += count;
    return `${openTag}${text}${closeTag}`;
  });
  // Fallback covers runs outside any w:p only (no paragraph context to join).
  // Earlier versions re-scanned every w:t here, so a counting pass reported
  // each in-paragraph hit twice while a redacting pass stayed quiet because
  // the marker text is inert. Splitting paragraph blocks from the gaps keeps
  // counting and redacting in agreement.
  const paraRe = new RegExp(P_RE.source, 'g');
  let out = '';
  let last = 0;
  let pm: RegExpExecArray | null;
  while ((pm = paraRe.exec(joined)) !== null) {
    const gap = joined.slice(last, pm.index);
    out += gap.replace(WT_RE, (_match, openTag: string, innerText: string, closeTag: string) => {
      // M2 fallback: runs outside any w:p take the body scanner (labeled
      // bare digits included).
      const { text, count } = scanAndRedactBodyText(innerText, redact);
      totalCount += count;
      return `${openTag}${text}${closeTag}`;
    });
    out += pm[0];
    last = pm.index + pm[0].length;
  }
  const tail = joined.slice(last);
  out += tail.replace(WT_RE, (_match, openTag: string, innerText: string, closeTag: string) => {
    const { text, count } = scanAndRedactBodyText(innerText, redact);
    totalCount += count;
    return `${openTag}${text}${closeTag}`;
  });
  return { xml: out, count: totalCount };
}

/** Joined-paragraph PII scan with span-accurate redaction across run slices. */
function redactSplitRunPiiInParagraph(inner: string, redact: boolean): { text: string; count: number } {
  if (!inner.includes('<w:t')) return { text: inner, count: 0 };
  const WT_RE = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;
  const segs: Array<{ open: string; inner: string; close: string }> = [];
  let m: RegExpExecArray | null;
  const re = new RegExp(WT_RE.source, 'g');
  while ((m = re.exec(inner)) !== null) segs.push({ open: m[1], inner: m[2], close: m[3] });
  if (segs.length === 0) return { text: inner, count: 0 };
  let count = 0;
  const PATS = [EMAIL_RE, PHONE_RE, SSN_RE, CREDIT_CARD_RE];
  for (const pat of PATS) {
    // Rebuild the joined text + offset map from current slice contents so
    // earlier-type redactions cannot shift or double-count later types.
    let joined = '';
    const segStart: number[] = [];
    const owner: number[] = [];
    segs.forEach((s, i) => {
      segStart.push(joined.length);
      for (let k = 0; k < s.inner.length; k++) owner.push(i);
      joined += s.inner;
    });
    if (joined.length === 0) continue;
    const rx = new RegExp(pat.source, 'g');
    const rawHits: Array<{ s: number; e: number }> = [];
    let mm: RegExpExecArray | null;
    while ((mm = rx.exec(joined)) !== null) {
      if (mm[0].length === 0) {
        rx.lastIndex++;
        continue;
      }
      rawHits.push({ s: mm.index, e: mm.index + mm[0].length });
    }
    // Clause-citation guard applies to the phone pattern only; other
    // patterns keep every hit.
    const hits =
      pat === PHONE_RE ? rawHits.filter((h) => !isClauseCitedNumber(joined, h.s, h.e)) : rawHits;
    count += hits.length;
    if (redact) {
      // Reverse order: higher-offset edits never disturb lower-offset mappings.
      for (let h = hits.length - 1; h >= 0; h--) {
        const { s, e } = hits[h];
        if (e <= s || s < 0 || e > joined.length) continue;
        const firstSeg = owner[s];
        const lastSeg = owner[e - 1];
        if (firstSeg === undefined || lastSeg === undefined) continue;
        const firstOff = s - segStart[firstSeg];
        if (firstSeg === lastSeg) {
          const seg = segs[firstSeg];
          const lastOff = e - segStart[firstSeg];
          seg.inner = seg.inner.slice(0, firstOff) + '[redacted]' + seg.inner.slice(lastOff);
        } else {
          const lastOff = e - segStart[lastSeg];
          segs[lastSeg].inner = segs[lastSeg].inner.slice(lastOff);
          for (let i = lastSeg - 1; i > firstSeg; i--) segs[i].inner = '';
          segs[firstSeg].inner = segs[firstSeg].inner.slice(0, firstOff) + '[redacted]';
        }
      }
    }
  }
  let idx = 0;
  const out = inner.replace(new RegExp(WT_RE.source, 'g'), () => {
    const s = segs[idx++];
    return `${s.open}${s.inner}${s.close}`;
  });
  // M2 labeled bare-digit pass (split-run aware, digits-only spans so the
  // label survives): rebuild from current slices, then redact digit spans.
  {
    let joined = '';
    const segStart: number[] = [];
    const owner: number[] = [];
    segs.forEach((s, i) => {
      segStart.push(joined.length);
      for (let k = 0; k < s.inner.length; k++) owner.push(i);
      joined += s.inner;
    });
    if (joined.length > 0) {
      const hits = labeledBareDigitSpans(joined);
      count += hits.length;
      if (redact) {
        for (let h = hits.length - 1; h >= 0; h--) {
          const { s, e } = hits[h];
          if (e <= s || s < 0 || e > joined.length) continue;
          const firstSeg = owner[s];
          const lastSeg = owner[e - 1];
          if (firstSeg === undefined || lastSeg === undefined) continue;
          const firstOff = s - segStart[firstSeg];
          if (firstSeg === lastSeg) {
            const seg = segs[firstSeg];
            const lastOff = e - segStart[firstSeg];
            seg.inner = seg.inner.slice(0, firstOff) + '[redacted]' + seg.inner.slice(lastOff);
          } else {
            const lastOff = e - segStart[lastSeg];
            segs[lastSeg].inner = segs[lastSeg].inner.slice(lastOff);
            for (let i = lastSeg - 1; i > firstSeg; i--) segs[i].inner = '';
            segs[firstSeg].inner = segs[firstSeg].inner.slice(0, firstOff) + '[redacted]';
          }
        }
      }
      if (redact && hits.length > 0) {
        let ridx = 0;
        return {
          text: inner.replace(new RegExp(WT_RE.source, 'g'), () => {
            const s = segs[ridx++];
            return `${s.open}${s.inner}${s.close}`;
          }),
          count,
        };
      }
    }
  }
  return { text: out, count };
}

export function sanitizeDocumentXml(
  xml: string,
  opts?: SanitizeOptions,
  partName = 'word XML part'
): {
  xml: string;
  removedRSIDs: number;
  purgedAuthors: number;
  acceptedRevisions: number;
  strippedCommentRefs: number;
  bodyPII: number;
} {
  assertNoForbiddenMarkup(xml, partName);
  const o = resolveOptions(opts);
  let out = xml;
  let removedRSIDs = 0;
  let purgedAuthors = 0;
  if (o.purgeRSIDs) {
    const hits = out.match(RSID_ATTR_RE);
    removedRSIDs += hits ? hits.length : 0;
    out = out.replace(RSID_ATTR_RE, '');
    // settings.xml rsid list (<w:rsid w:val="…"/> + <w:rsidRoot w:val="…"/> elements).
    const eHits = out.match(RSID_ELEM_RE);
    removedRSIDs += eHits ? eHits.length : 0;
    out = out.replace(RSID_ELEM_RE, '');
    // Strip the now-empty <w:rsids></w:rsids> wrapper cleanly (incl. whitespace).
    out = out.replace(RSIDS_EMPTY_RE, '');
    out = out.replace(RSIDS_SELF_CLOSING_RE, '');
  }
  if (o.purgeAuthors) {
    const replacement = escapeXmlAttr(o.replacementAuthor);
    const aHits = out.match(AUTHOR_ATTR_RE);
    purgedAuthors += aHits ? aHits.length : 0;
    // F3: preserve the matched prefix (w: vs w15:) via replace callback so
    // word/people.xml keeps w15:author (never rewrite to w:author).
    out = out.replace(AUTHOR_ATTR_RE, (_m: string, prefix: string) => ` ${prefix}:author="${replacement}"`);
    const iHits = out.match(INITIALS_ATTR_RE);
    purgedAuthors += iHits ? iHits.length : 0;
    const initials = escapeXmlAttr(
      o.replacementAuthor
        .split(/[\s._-]+/)
        .filter(Boolean)
        .map((s) => s[0])
        .join('')
        .slice(0, 4)
        .toUpperCase() || 'A'
    );
    out = out.replace(INITIALS_ATTR_RE, ` w:initials="${initials}"`);
  }
  let acceptedRevisions = 0;
  let strippedCommentRefs = 0;
  if (o.acceptAllRevisions || o.stripComments) {
    const r = applyRevisionAndCommentPass(out, o, partName);
    out = r.xml;
    acceptedRevisions = r.acceptedRevisions;
    strippedCommentRefs = r.strippedCommentRefs;
  }
  let bodyPII = 0;
  if (o.purgePII) {
    if (isBodyPartName(partName) || isCommentFamilyPartName(partName)) {
      // w:t prose (split-run aware) first; markers are PII-inert so the
      // follow-up passes never double-count.
      const piiRes = redactBodyPiiInWordText(out, true);
      out = piiRes.xml;
      bodyPII += piiRes.count;
      // Non-w:t carriers: field-code text (including w:instrText), delText,
      // comment metadata text.
      const textRes = redactTextNodePii(out, true);
      out = textRes.xml;
      bodyPII += textRes.count;
      if (isCommentFamilyPartName(partName)) {
        // people.xml w15:userId and comment carrier attributes.
        const attrRes = redactAttributePii(out, true);
        out = attrRes.xml;
        bodyPII += attrRes.count;
      }
      if (isBodyPartName(partName)) {
        // Drawing alternative text: descr/title/name on wp:docPr and
        // pic:cNvPr companions. Targeted to those attributes only so
        // structural ids and geometry stay intact.
        const altRes = redactDrawingAltPii(out, true);
        out = altRes.xml;
        bodyPII += altRes.count;
      }
    } else if (isGenericTextPiiPartName(partName)) {
      // charts (a:t/c:v), customXml items, style-family element text.
      // Style-family parts additionally scan name attributes (w:name/w:val
      // carry the prose there). Chart and customXml parts scan attribute
      // values with the attribute scanner (labeled-only bare digits, so
      // item ids and numeric format attrs without a label stay intact).
      const textRes = redactTextNodePii(out, true);
      out = textRes.xml;
      bodyPII += textRes.count;
      if (isStyleFamilyPartName(partName)) {
        const attrRes = redactAttributePii(out, true);
        out = attrRes.xml;
        bodyPII += attrRes.count;
      } else if (isChartOrCustomXmlPartName(partName)) {
        const attrRes = redactAttributePii(out, true);
        out = attrRes.xml;
        bodyPII += attrRes.count;
      }
    }
  }
  return { xml: out, removedRSIDs, purgedAuthors, acceptedRevisions, strippedCommentRefs, bodyPII };
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

// Comment part graph. Relationship types + content types mirror
// ooxml/comments.ts (duplicated here so the sanitizer stays dependency-free).
const COMMENT_REL_TYPES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  'http://schemas.microsoft.com/office/2011/relationships/people',
]);
const COMMENT_PART_NAMES = ['word/comments.xml', 'word/commentsExtended.xml', 'word/people.xml'];
const COMMENT_PART_CTS = new Map([
  ['/word/comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'],
  [
    '/word/commentsExtended.xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
  ],
  ['/word/people.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml'],
]);
const COMMENT_REF_LOCALS = ['commentRangeStart', 'commentRangeEnd', 'commentReference'];

/** Fail-closed DOM parse: throws E_VALIDATION on malformed XML (never silent clean). */
function parseFailClosedXml(body: string, context: string): Document {
  assertNoForbiddenMarkup(body, context);
  let firstMsg = '';
  let sawIssue = false;
  try {
    const doc = new DOMParser({
      errorHandler: {
        warning(msg: string) {
          sawIssue = true;
          if (!firstMsg) firstMsg = String(msg).slice(0, 200);
        },
        error(msg: string) {
          sawIssue = true;
          if (!firstMsg) firstMsg = String(msg).slice(0, 200);
        },
        fatalError(msg: string) {
          sawIssue = true;
          if (!firstMsg) firstMsg = String(msg).slice(0, 200);
        },
      },
    }).parseFromString(body, 'text/xml');
    if (sawIssue) {
      throw new Error(
        `E_VALIDATION: malformed XML in ${context} cannot be safely sanitized (${firstMsg}). ` +
          `Remedy: repair the part XML or re-save the document in Word before sanitizing.`
      );
    }
    if (!doc || !doc.documentElement) {
      throw new Error(
        `E_VALIDATION: malformed XML in ${context} cannot be safely sanitized (unparseable). ` +
          `Remedy: repair the part XML or re-save the document in Word before sanitizing.`
      );
    }
    const perrs = doc.getElementsByTagName('parsererror');
    if (perrs && perrs.length > 0) {
      throw new Error(
        `E_VALIDATION: malformed XML in ${context} cannot be safely sanitized (parsererror). ` +
          `Remedy: repair the part XML or re-save the document in Word before sanitizing.`
      );
    }
    return doc;
  } catch (err) {
    if (err instanceof Error && err.message.includes('E_VALIDATION')) throw err;
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(
      `E_VALIDATION: malformed XML in ${context} cannot be safely sanitized (${msg.slice(0, 200)}). ` +
        `Remedy: repair the part XML or re-save the document in Word before sanitizing.`
    );
  }
}

function splitXmlProlog(xml: string): { head: string; body: string } {
  const m = /^(\s*<\?xml[^?]*\?>\s*)/.exec(xml);
  if (!m) return { head: '', body: xml };
  return { head: m[1], body: xml.slice(m[1].length) };
}

function serializeWithProlog(head: string, doc: Document): string {
  return head + new XMLSerializer().serializeToString(doc);
}

function firstEl(doc: Document, ns: string, local: string): Element | null {
  const list = doc.getElementsByTagNameNS(ns, local);
  return list.length > 0 ? (list.item(0) as Element | null) : null;
}

/**
 * Accepts tracked changes in a parsed word part (Word "accept all" semantics).
 * Drops `w:del`/`w:moveFrom` subtrees entirely (deleted/moved-away text gone,
 * incl. any nested `w:ins`/`w:moveTo`), then unwraps `w:ins`/`w:moveTo`
 * keeping children. Drops move range markers, resolves cell revision markers
 * (`cellIns`/`cellMerge` marker dropped keeping the cell, `cellDel` drops its
 * containing `w:tc`), and drops the `*PrChange` formatting-history wrappers
 * keeping current formatting. Namespace-aware (prefix agnostic) wherever the
 * nodes sit — paragraphs, tables, headers/footers. Live-query loop: the first
 * match in document order is always an outermost remaining node, so nested
 * pairs resolve correctly.
 */
function acceptAllRevisionsInDom(doc: Document): number {
  let count = 0;
  const bump = (): boolean => {
    count++;
    return count > 1000000;
  };
  let el: Element | null;
  // Drops first (deleted/moved-away content gone, incl. nested inserts inside).
  for (const local of ['moveFrom', 'del']) {
    while ((el = firstEl(doc, W_NS, local)) !== null) {
      el.parentNode?.removeChild(el);
      if (bump()) return count;
    }
  }
  // Unwraps (keep destination/inserted content).
  for (const local of ['moveTo', 'ins']) {
    while ((el = firstEl(doc, W_NS, local)) !== null) {
      const parent = el.parentNode;
      if (!parent) break;
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
      if (bump()) return count;
    }
  }
  // Move/customXml range markers: empty anchors carrying author/date — drop all.
  for (const local of [
    'moveFromRangeStart',
    'moveFromRangeEnd',
    'moveToRangeStart',
    'moveToRangeEnd',
    'customXmlInsRangeStart',
    'customXmlInsRangeEnd',
    'customXmlDelRangeStart',
    'customXmlDelRangeEnd',
    'customXmlMoveFromRangeStart',
    'customXmlMoveFromRangeEnd',
    'customXmlMoveToRangeStart',
    'customXmlMoveToRangeEnd',
  ]) {
    while ((el = firstEl(doc, W_NS, local)) !== null) {
      el.parentNode?.removeChild(el);
      if (bump()) return count;
    }
  }
  // Deleted cells: accepting removes the whole containing w:tc (its text is
  // deleted content). Runs before cellIns/cellMerge so a doomed cell takes
  // any co-located insert marker with it (single count for the cell).
  while ((el = firstEl(doc, W_NS, 'cellDel')) !== null) {
    const tc = ancestorInNs(el, W_NS, 'tc');
    if (tc?.parentNode) tc.parentNode.removeChild(tc);
    else el.parentNode?.removeChild(el);
    if (bump()) return count;
  }
  // Inserted/merged cells: accepting keeps the cell, drops the marker only.
  for (const local of ['cellIns', 'cellMerge']) {
    while ((el = firstEl(doc, W_NS, local)) !== null) {
      el.parentNode?.removeChild(el);
      if (bump()) return count;
    }
  }
  // Formatting-history wrappers: the sibling w:rPr/w:pPr holds accepted
  // formatting; the *Change child holds the previous value — drop the wrapper.
  for (const local of [
    'rPrChange',
    'pPrChange',
    'sectPrChange',
    'numberingChange',
    'tblPrChange',
    'trPrChange',
    'tcPrChange',
    'tPrChange',
  ]) {
    while ((el = firstEl(doc, W_NS, local)) !== null) {
      el.parentNode?.removeChild(el);
      if (bump()) return count;
    }
  }
  // Future-proof: any other W:*Change revision wrapper (accept = drop wrapper).
  for (;;) {
    const change = firstWChangeElement(doc);
    if (!change) break;
    change.parentNode?.removeChild(change);
    if (bump()) return count;
  }
  return count;
}

/** Walks up to the nearest ancestor with the given namespace + local name. */
function ancestorInNs(el: Element, ns: string, local: string): Element | null {
  let node: Node | null = el.parentNode;
  while (node) {
    if (node.nodeType === 1) {
      const e = node as Element;
      const nsMatch = (e.namespaceURI ?? '') === ns;
      const localMatch =
        (e.localName ?? '') === local ||
        e.tagName === local ||
        e.tagName.endsWith(`:${local}`);
      if (nsMatch && localMatch) return e;
    }
    node = node.parentNode;
  }
  return null;
}

/** Finds the first remaining W-namespace element whose local ends in "Change". */
function firstWChangeElement(doc: Document): Element | null {
  const all = doc.getElementsByTagNameNS(W_NS, '*');
  for (let i = 0; i < all.length; i++) {
    const e = all.item(i) as Element | null;
    if (!e) continue;
    const local = e.localName ?? e.tagName.split(':').pop() ?? '';
    if (local.endsWith('Change')) return e;
  }
  return null;
}

/** Removes all comment reference elements (range start/end + reference runs). */
function stripCommentRefsInDom(doc: Document): number {
  let count = 0;
  for (const local of COMMENT_REF_LOCALS) {
    let el: Element | null;
    while ((el = firstEl(doc, W_NS, local)) !== null) {
      el.parentNode?.removeChild(el);
      count++;
      if (count > 1000000) break;
    }
  }
  return count;
}

/**
 * Shared revision/comment pass for every word XML part (document, headers,
 * footers, styles, numbering, …). Returns the input string untouched
 * (byte-identical) when there is nothing to change. Fail-closed: malformed
 * XML under an active flag throws E_VALIDATION (never a silent clean report).
 */
function applyRevisionAndCommentPass(
  xml: string,
  o: { acceptAllRevisions: boolean; stripComments: boolean },
  context = 'word XML part'
): { xml: string; acceptedRevisions: number; strippedCommentRefs: number } {
  assertNoForbiddenMarkup(xml, context);
  const none = { xml, acceptedRevisions: 0, strippedCommentRefs: 0 };
  const { head, body } = splitXmlProlog(xml);
  const doc = parseFailClosedXml(body, context);
  let acceptedRevisions = 0;
  let strippedCommentRefs = 0;
  if (o.acceptAllRevisions) acceptedRevisions = acceptAllRevisionsInDom(doc);
  if (o.stripComments) strippedCommentRefs = stripCommentRefsInDom(doc);
  if (acceptedRevisions === 0 && strippedCommentRefs === 0) return none;
  return {
    xml: serializeWithProlog(head, doc),
    acceptedRevisions,
    strippedCommentRefs,
  };
}

/**
 * Removes `Override` elements for dropped comment parts from
 * `[Content_Types].xml`. Returns input untouched when nothing matches.
 */
function stripCommentContentTypes(xml: string, droppedPartNames: Set<string>): { xml: string; removed: number } {
  assertNoForbiddenMarkup(xml, '[Content_Types].xml');
  const none = { xml, removed: 0 };
  if (droppedPartNames.size === 0 || !xml.includes('Override')) return none;
  const targets = new Set<string>();
  for (const name of droppedPartNames) targets.add(`/${name}`);
  let hit = false;
  for (const t of targets) {
    if (xml.includes(t)) {
      hit = true;
      break;
    }
  }
  if (!hit) return none;
  const { head, body } = splitXmlProlog(xml);
  const doc = parseFailClosedXml(body, '[Content_Types].xml');
  const list = doc.getElementsByTagNameNS(CT_NS, 'Override');
  const doomed: Element[] = [];
  for (let i = 0; i < list.length; i++) {
    const el = list.item(i) as Element;
    const partName = el.getAttribute('PartName');
    if (partName && targets.has(partName) && COMMENT_PART_CTS.has(partName)) doomed.push(el);
  }
  if (doomed.length === 0) return none;
  for (const el of doomed) el.parentNode?.removeChild(el);
  return { xml: serializeWithProlog(head, doc), removed: doomed.length };
}

/**
 * Removes comment `Relationship` entries (comments/commentsExtended/people
 * types whose target is a dropped part) from a `.rels` part. Both the type
 * AND the target must match, so unrelated rels are never touched.
 */
function stripCommentRelationships(
  xml: string,
  droppedPartNames: Set<string>,
  context = 'word/_rels/*.rels'
): { xml: string; removed: number } {
  assertNoForbiddenMarkup(xml, context);
  const none = { xml, removed: 0 };
  if (droppedPartNames.size === 0 || !xml.includes('Relationship')) return none;
  const basenames = new Set<string>();
  for (const name of droppedPartNames) {
    const base = name.split('/').pop();
    if (base) basenames.add(base);
  }
  const { head, body } = splitXmlProlog(xml);
  const doc = parseFailClosedXml(body, context);
  const list = doc.getElementsByTagNameNS(REL_NS, 'Relationship');
  const doomed: Element[] = [];
  for (let i = 0; i < list.length; i++) {
    const el = list.item(i) as Element;
    const type = el.getAttribute('Type');
    const target = el.getAttribute('Target');
    if (!type || !target || !COMMENT_REL_TYPES.has(type)) continue;
    const base = target.split('/').pop();
    if (base && basenames.has(base)) doomed.push(el);
  }
  if (doomed.length === 0) return none;
  for (const el of doomed) el.parentNode?.removeChild(el);
  return { xml: serializeWithProlog(head, doc), removed: doomed.length };
}

function replaceTagText(xml: string, tag: string, replacement: string): { xml: string; count: number } {
  const re = new RegExp(`(<${tag}[^>]*>)([\\s\\S]*?)(</${tag}>)`, 'g');
  let count = 0;
  const out = xml.replace(re, (_m, open: string, _inner: string, close: string) => {
    count++;
    return `${open}${replacement}${close}`;
  });
  return { xml: out, count };
}

export function sanitizeCorePropertiesXml(
  xml: string,
  opts?: SanitizeOptions
): { xml: string; purgedAuthors: number; purgedPII: number } {
  assertNoForbiddenMarkup(xml, 'docProps/core.xml');
  const o = resolveOptions(opts);
  let out = xml;
  let purgedAuthors = 0;
  let purgedPII = 0;
  // PII first (on the original values) so emails/phones inside author fields
  // are counted before the author replacement overwrites them; then authors.
  if (o.purgePII) {
    const emailRe = new RegExp(EMAIL_RE.source, 'g');
    const emails = out.match(emailRe);
    purgedPII += emails ? emails.length : 0;
    out = out.replace(emailRe, '[redacted]');
    const guardedPhones = redactPhonesWithClauseGuard(out, true);
    purgedPII += guardedPhones.count;
    out = guardedPhones.text;
    const ssnRe = new RegExp(SSN_RE.source, 'g');
    const ssns = out.match(ssnRe);
    purgedPII += ssns ? ssns.length : 0;
    out = out.replace(ssnRe, '[redacted]');
    const cardRe = new RegExp(CREDIT_CARD_RE.source, 'g');
    const ccs = out.match(cardRe);
    purgedPII += ccs ? ccs.length : 0;
    out = out.replace(cardRe, '[redacted]');
  }
  if (o.purgeAuthors) {
    const replacement = escapeXmlAttr(o.replacementAuthor);
    if (o.zeroRetention) {
      // Zero-retention: empty authorship instead of stamping a replacement
      // author — no attribution string survives. Counts stay honest.
      for (const tag of ['dc:creator', 'cp:lastModifiedBy']) {
        const r = replaceTagText(out, tag, '');
        out = r.xml;
        purgedAuthors += r.count;
      }
    } else {
      for (const tag of ['dc:creator', 'cp:lastModifiedBy']) {
        const r = replaceTagText(out, tag, replacement);
        out = r.xml;
        purgedAuthors += r.count;
      }
    }
  }
  // P0-7: In docProps/core.xml: Under zeroRetention: true (or when purging metadata), purge dc:title, dc:subject, cp:keywords, dc:description, cp:category.
  // Zero-retention additionally empties revision stamps (dcterms:created,
  // dcterms:modified, cp:revision) — timestamps and revision counts identify
  // editing sessions and must not survive a zero-retention export.
  if (o.zeroRetention) {
    for (const tag of [
      'dc:title',
      'dc:subject',
      'cp:keywords',
      'dc:description',
      'cp:category',
      'dcterms:created',
      'dcterms:modified',
      'cp:revision',
    ]) {
      const r = replaceTagText(out, tag, '');
      out = r.xml;
      purgedAuthors += r.count;
    }
  }
  return { xml: out, purgedAuthors, purgedPII };
}

export function sanitizeAppPropertiesXml(
  xml: string,
  opts?: SanitizeOptions
): { xml: string; purgedAuthors: number; purgedPII: number } {
  assertNoForbiddenMarkup(xml, 'docProps/app.xml');
  const o = resolveOptions(opts);
  let out = xml;
  let purgedAuthors = 0;
  let purgedPII = 0;
  // Same ordering as core props: PII on originals first, then authors.
  if (o.purgePII) {
    const emailRe = new RegExp(EMAIL_RE.source, 'g');
    const emails = out.match(emailRe);
    purgedPII += emails ? emails.length : 0;
    out = out.replace(emailRe, '[redacted]');
    const guardedPhones = redactPhonesWithClauseGuard(out, true);
    purgedPII += guardedPhones.count;
    out = guardedPhones.text;
    const ssnRe = new RegExp(SSN_RE.source, 'g');
    const ssns = out.match(ssnRe);
    purgedPII += ssns ? ssns.length : 0;
    out = out.replace(ssnRe, '[redacted]');
    const cardRe = new RegExp(CREDIT_CARD_RE.source, 'g');
    const ccs = out.match(cardRe);
    purgedPII += ccs ? ccs.length : 0;
    out = out.replace(cardRe, '[redacted]');
  }
  if (o.purgeAuthors) {
    const replacement = escapeXmlAttr(o.replacementAuthor);
    for (const tag of ['Company', 'Manager']) {
      const r = replaceTagText(out, tag, replacement);
      out = r.xml;
      purgedAuthors += r.count;
    }
  }
  return { xml: out, purgedAuthors, purgedPII };
}

/**
 * Redacts PII inside docProps/custom.xml string values.
 *
 * Custom properties are user-defined (<property name="…"><vt:lpwstr>…</vt:lpwstr>)
 * so there are no fixed author tags to overwrite — the purge is PII-only
 * (email / phone / SSN / credit-card patterns → '[redacted]') applied across the part. Returned
 * purgedAuthors is always 0 (kept for report-shape consistency with core/app).
 */
export function sanitizeCustomPropertiesXml(
  xml: string,
  opts?: SanitizeOptions
): { xml: string; purgedAuthors: number; purgedPII: number } {
  assertNoForbiddenMarkup(xml, 'docProps/custom.xml');
  const o = resolveOptions(opts);
  let out = xml;
  let purgedPII = 0;
  if (o.purgePII) {
    const emailRe = new RegExp(EMAIL_RE.source, 'g');
    const emails = out.match(emailRe);
    purgedPII += emails ? emails.length : 0;
    out = out.replace(emailRe, '[redacted]');
    const guardedPhones = redactPhonesWithClauseGuard(out, true);
    purgedPII += guardedPhones.count;
    out = guardedPhones.text;
    const ssnRe = new RegExp(SSN_RE.source, 'g');
    const ssns = out.match(ssnRe);
    purgedPII += ssns ? ssns.length : 0;
    out = out.replace(ssnRe, '[redacted]');
    const cardRe = new RegExp(CREDIT_CARD_RE.source, 'g');
    const ccs = out.match(cardRe);
    purgedPII += ccs ? ccs.length : 0;
    out = out.replace(cardRe, '[redacted]');
  }
  return { xml: out, purgedAuthors: 0, purgedPII };
}

/**
 * Strips JPEG EXIF/XMP (APP1) and Photoshop IRB (APP13) segments.
 * Walks markers without decoding scan data; returns the original buffer
 * untouched when it is not a JPEG. Bounded: single linear pass, O(n) time,
 * O(1) extra segments retained.
 */
export function stripJpegExif(buf: Buffer): { buf: Buffer; strippedBytes: number } {
  if (buf.length < 4 || buf[0] !== 0xff || buf[1] !== 0xd8) return { buf, strippedBytes: 0 };
  const out: Buffer[] = [buf.subarray(0, 2)];
  let strippedBytes = 0;
  let pos = 2;
  const isExifApp1 = (payload: Buffer): boolean => {
    if (payload.length >= 6 && payload.subarray(0, 6).toString('binary') === 'Exif\0\0') return true;
    const head = payload.subarray(0, 64).toString('binary');
    return head.includes('http://ns.adobe.com/xap/') || head.includes('http://ns.adobe');
  };
  while (pos + 4 <= buf.length) {
    if (buf[pos] !== 0xff) break;
    const marker = buf[pos + 1];
    // Standalone markers (no length field).
    if (marker === 0xd8 || marker === 0xd9 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      out.push(buf.subarray(pos, pos + 2));
      pos += 2;
      if (marker === 0xd9) break;
      continue;
    }
    if (pos + 4 > buf.length) break;
    const segLen = buf.readUInt16BE(pos + 2);
    if (segLen < 2 || pos + 2 + segLen > buf.length) break;
    const segment = buf.subarray(pos, pos + 2 + segLen);
    const payload = buf.subarray(pos + 4, pos + 2 + segLen);
    const isApp1 = marker === 0xe1;
    const isApp13 = marker === 0xed;
    if ((isApp1 && isExifApp1(payload)) || isApp13) {
      // APP1: drop Exif/XMP only (both share the APP1 marker and are PII
      // carriers). Non-Exif APP1 segments are preserved. APP13: Photoshop
      // IRB (captions/authors). ICC (APP2) kept.
      strippedBytes += segment.length;
      pos += 2 + segLen;
      continue;
    }
    out.push(segment);
    pos += 2 + segLen;
    // Start-of-scan: remainder is entropy-coded image data — copy verbatim.
    if (marker === 0xda) {
      out.push(buf.subarray(pos));
      break;
    }
  }
  if (strippedBytes === 0) return { buf, strippedBytes: 0 };
  return { buf: Buffer.concat(out), strippedBytes };
}

const PNG_TEXT_KEYWORDS = new Set([
  'author',
  'creator',
  'description',
  'comment',
  'software',
  'title',
  'copyright',
  'creation time',
  'disclaimer',
  'warning',
  'source',
  'artist',
  // Adobe XMP carriers (matched case-insensitively via toLowerCase lookup):
  // iTXt keyword 'XML:com.adobe.xmp' (Photoshop) and bare 'xmp'.
  'xml:com.adobe.xmp',
  'xmp',
]);

/**
 * Strips PNG textual PII chunks (tEXt/zTXt/iTXt) whose keyword is an
 * authorship/description carrier, plus the binary eXIf chunk (PNG spec v1.5.0
 * EXIF/GPS carrier from cameras/phones). IHDR/PLTE/IDAT/IEND and color chunks
 * (gAMA/cHRM/sRGB/iCCP) are preserved. Returns original when not a PNG.
 * Keyword match is case-insensitive (lowercased lookup).
 */
export function stripPngTextChunks(buf: Buffer): { buf: Buffer; strippedBytes: number } {
  const PNG_SIG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  if (buf.length < 8 || !buf.subarray(0, 8).equals(PNG_SIG)) return { buf, strippedBytes: 0 };
  const out: Buffer[] = [buf.subarray(0, 8)];
  let strippedBytes = 0;
  let pos = 8;
  while (pos + 8 <= buf.length) {
    const len = buf.readUInt32BE(pos);
    const type = buf.subarray(pos + 4, pos + 8).toString('ascii');
    const total = 12 + len;
    if (pos + total > buf.length) break;
    const chunk = buf.subarray(pos, pos + total);
    // eXIf: binary EXIF/GPS — strip unconditionally (always a PII carrier).
    if (type === 'eXIf') {
      strippedBytes += chunk.length;
      pos += total;
      continue;
    }
    if (type === 'tEXt' || type === 'zTXt' || type === 'iTXt') {
      const nul = chunk.indexOf(0x00, 8);
      const keyword =
        nul > 8 ? chunk.subarray(8, nul).toString('latin1').trim().toLowerCase() : '';
      if (PNG_TEXT_KEYWORDS.has(keyword)) {
        strippedBytes += chunk.length;
        pos += total;
        continue;
      }
    }
    out.push(chunk);
    pos += total;
    if (type === 'IEND') break;
  }
  if (strippedBytes === 0) return { buf, strippedBytes: 0 };
  // Copy any trailing bytes verbatim (defensive; well-formed PNGs end at IEND).
  if (pos < buf.length) out.push(buf.subarray(pos));
  return { buf: Buffer.concat(out), strippedBytes };
}

export function stripImageExif(buf: Buffer, fileName: string): { buf: Buffer; strippedBytes: number } {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return stripJpegExif(buf);
  if (lower.endsWith('.png')) return stripPngTextChunks(buf);
  return { buf, strippedBytes: 0 };
}

/**
 * Redacts contact identifiers inside an embedded SpreadsheetML workbook
 * (`word/embeddings/*.xlsx`, itself a zip). Every inner XML-ish entry is
 * gated for forbidden markup (fail-closed) then scanned for element-text and
 * attribute PII with the shared patterns; opaque inner blobs pass through
 * untouched. Returns the input buffer verbatim when nothing redacts, so
 * clean workbooks stay byte-identical. Non-zip payloads pass through
 * untouched (opaque bins stay opaque).
 */
export function sanitizeEmbeddedWorkbookBytes(
  buf: Buffer,
  outerName: string
): { buf: Buffer; purgedPII: number } {
  const none = { buf, purgedPII: 0 };
  if (!buf || buf.length < 4) return none;
  let inner: AdmZip;
  try {
    inner = new AdmZip(buf);
  } catch {
    return none;
  }
  let entries: ReturnType<AdmZip['getEntries']>;
  try {
    entries = inner.getEntries();
  } catch {
    return none;
  }
  if (!Array.isArray(entries) || entries.length === 0) return none;
  let purgedPII = 0;
  let changed = false;
  for (const entry of entries) {
    if (!entry || entry.isDirectory) continue;
    const innerName = entry.entryName;
    let data: Buffer;
    try {
      data = Buffer.from(entry.getData());
    } catch {
      continue;
    }
    if (!isXmlPart(innerName)) continue;
    const xml = data.toString('utf8');
    assertNoForbiddenMarkup(xml, `${outerName}!${innerName}`);
    const textRes = redactTextNodePii(xml, true);
    const attrRes = redactAttributePii(textRes.xml, true);
    purgedPII += textRes.count + attrRes.count;
    if (attrRes.xml !== xml) {
      changed = true;
      inner.updateFile(innerName, Buffer.from(attrRes.xml, 'utf8'));
    }
  }
  if (!changed) return { buf, purgedPII };
  clampZipEntryTimes(inner);
  return { buf: inner.toBuffer(), purgedPII };
}

const XML_PART_RE =
  /^(word\/document\.xml|word\/styles\.xml|word\/stylesWithEffects\.xml|word\/numbering\.xml|word\/comments\.xml|word\/commentsExtended\.xml|word\/footnotes\.xml|word\/endnotes\.xml|word\/settings\.xml|word\/header.*\.xml|word\/footer.*\.xml|word\/people\.xml|word\/customXml\/.*\.xml|word\/charts\/.*\.xml|word\/embeddings\/.*\.xml)$/i;

/**
 * Sanitizes an in-memory package map (entry path -> bytes). Word XML parts
 * (document/styles/stylesWithEffects/numbering/comments/commentsExtended/footnotes/endnotes/
 * settings/headers/footers/people/customXml/charts/embeddings XML) are purged of
 * RSIDs/authors, embedded workbook packages (`word/embeddings/*.xlsx`) have
 * their inner sheet text scanned for contact identifiers, core/app/custom
 * props of authors/PII, images of EXIF.
 * With `acceptAllRevisions`, tracked deletions/moves are dropped, insertions/
 * move-destinations unwrapped, cell markers resolved, and `*PrChange`
 * wrappers dropped in every word part; with `stripComments`, the comment parts
 * are deleted (adjuncts gated on `word/comments.xml` presence, orphans incl.)
 * and their content-type / relationship / reference traces removed.
 * The input map is never mutated; a fresh map is returned with the report.
 */
export function sanitizePackageEntries(
  files: Map<string, Buffer>,
  opts?: SanitizeOptions
): { files: Map<string, Buffer>; report: SanitizeReport } {
  const o = resolveOptions(opts);
  // Macro/active-content gate runs FIRST (fail-closed before any other
  // work): default refuses naming the parts, allow preserves them verbatim,
  // strip removes them with per-entry accounting below.
  let current = files;
  let macroParts: string[] = [];
  let strippedMacros: string[] = [];
  const macroTouched: string[] = [];
  {
    const readText = (name: string): string | null => {
      const buf = current.get(name);
      return buf ? buf.toString('utf8') : null;
    };
    macroParts = findMacroPartNames(current.keys(), readText);
    if (macroParts.length > 0) {
      if (o.allowMacros) {
        // Preserve: bins and macro rels flow through the verbatim branches
        // below untouched (byte-identical).
      } else if (o.stripMacros) {
        const working = new Map(current);
        const stripped = applyMacroStrip(mapMacroStore(working));
        current = working;
        strippedMacros = stripped.removed;
        macroTouched.push(...stripped.removed, ...stripped.edited);
      } else {
        throw macroRefusalError(macroParts, 'sanitize');
      }
    }
  }

  // Ingress XXE and rogue entry checks
  const out = new Map<string, Buffer>();
  const touched = new Set<string>();
  let removedRSIDs = 0;
  let purgedAuthors = 0;
  let strippedExifBytes = 0;
  let propsPII = 0;
  let bodyPII = 0;
  let acceptedRevisions = 0;
  let strippedComments = 0;
  const strippedSvg: string[] = [];
  const droppedRogue: string[] = [];
  // Per-part scan evidence (see SanitizeReport.scannedParts): every entry
  // that passes through a content scan below is recorded, whether or not
  // anything was purged, so all-zero purge counts stay trustworthy.
  const scanned = new Set<string>();
  for (const n of macroTouched) touched.add(n);
  const noteDropped = (name: string): void => {
    if (!droppedRogue.includes(name)) {
      droppedRogue.push(name);
      touched.add(name);
    }
  };
  for (const [name, data] of current) {
    if (isDroppedEntry(name)) {
      noteDropped(name);
      continue;
    }
    if (isXmlPart(name)) {
      assertNoForbiddenMarkup(data.toString('utf8'), name);
    }
  }

  // D-2: drop comment parts before the per-part pass so their bodies
  // (author names, privileged text) leave the package entirely.
  // Choice (documented): adjuncts (commentsExtended/people) drop only when
  // word/comments.xml is present. Standalone people.xml from external Word
  // co-authoring without comments is preserved (not a comment adjunct in that
  // shape) to avoid over-deletion; kept adjuncts still get RSID/author
  // purging via the shared XML pass. Orphan word/_rels/<part>.rels for each
  // dropped part is removed too (no dangling part left behind).
  const droppedCommentParts = new Set<string>();
  const droppedOrphanRels = new Set<string>();
  if (o.stripComments) {
    const hasComments = current.has('word/comments.xml');
    for (const name of COMMENT_PART_NAMES) {
      if (!current.has(name)) continue;
      if ((name === 'word/people.xml' || name === 'word/commentsExtended.xml') && !hasComments) {
        continue;
      }
      droppedCommentParts.add(name);
      touched.add(name);
    }
    for (const name of droppedCommentParts) {
      const base = name.split('/').pop();
      if (!base) continue;
      const relsName = `word/_rels/${base}.rels`;
      if (current.has(relsName)) {
        droppedOrphanRels.add(relsName);
        touched.add(relsName);
      }
    }
    strippedComments += droppedCommentParts.size + droppedOrphanRels.size;
  }

  for (const [name, data] of current) {
    if (droppedCommentParts.has(name) || droppedOrphanRels.has(name)) continue;
    if (isDroppedEntry(name)) {
      noteDropped(name);
      continue;
    }
    if (XML_PART_RE.test(name)) {
      scanned.add(name);
      const xml = data.toString('utf8');
      const r = sanitizeDocumentXml(xml, o, name);
      removedRSIDs += r.removedRSIDs;
      purgedAuthors += r.purgedAuthors;
      acceptedRevisions += r.acceptedRevisions;
      strippedComments += r.strippedCommentRefs;
      bodyPII += r.bodyPII;
      if (r.xml !== xml) touched.add(name);
      out.set(name, Buffer.from(r.xml, 'utf8'));
    } else if (name === '[Content_Types].xml' && droppedCommentParts.size > 0) {
      scanned.add(name);
      const xml = data.toString('utf8');
      const r = stripCommentContentTypes(xml, droppedCommentParts);
      strippedComments += r.removed;
      if (r.xml !== xml) touched.add(name);
      out.set(name, Buffer.from(r.xml, 'utf8'));
    } else if (name.endsWith('.rels') && droppedCommentParts.size > 0) {
      scanned.add(name);
      const xml = data.toString('utf8');
      const r = stripCommentRelationships(xml, droppedCommentParts, name);
      strippedComments += r.removed;
      if (r.xml !== xml) touched.add(name);
      out.set(name, Buffer.from(r.xml, 'utf8'));
    } else if (name === 'docProps/core.xml') {
      scanned.add(name);
      const xml = data.toString('utf8');
      const r = sanitizeCorePropertiesXml(xml, o);
      purgedAuthors += r.purgedAuthors;
      propsPII += r.purgedPII;
      if (r.xml !== xml) touched.add(name);
      out.set(name, Buffer.from(r.xml, 'utf8'));
    } else if (name === 'docProps/app.xml') {
      scanned.add(name);
      const xml = data.toString('utf8');
      const r = sanitizeAppPropertiesXml(xml, o);
      purgedAuthors += r.purgedAuthors;
      propsPII += r.purgedPII;
      if (r.xml !== xml) touched.add(name);
      out.set(name, Buffer.from(r.xml, 'utf8'));
    } else if (name === 'docProps/custom.xml') {
      scanned.add(name);
      const xml = data.toString('utf8');
      const r = sanitizeCustomPropertiesXml(xml, o);
      purgedAuthors += r.purgedAuthors;
      propsPII += r.purgedPII;
      if (r.xml !== xml) touched.add(name);
      out.set(name, Buffer.from(r.xml, 'utf8'));
    } else if (isEmbeddingWorkbookPartName(name) && o.purgePII) {
      scanned.add(name);
      const r = sanitizeEmbeddedWorkbookBytes(Buffer.from(data), name);
      bodyPII += r.purgedPII;
      if (!r.buf.equals(Buffer.from(data))) touched.add(name);
      out.set(name, r.buf);
    } else if (name.startsWith('word/media/') && o.purgeExif) {
      scanned.add(name);
      if (/\.svg$/i.test(name)) {
        // D5 sanitize path (strip-loud): script-capable SVG constructs are
        // stripped (never refused here) and the part is named loudly below.
        // Untouched bytes flow through verbatim so clean SVGs stay
        // byte-identical (no UTF-8 round-trip when nothing stripped).
        const r = stripSvgActiveContent(data.toString('utf8'));
        if (r.removed.length > 0) {
          touched.add(name);
          strippedSvg.push(name);
          out.set(name, Buffer.from(r.cleaned, 'utf8'));
        } else {
          out.set(name, Buffer.from(data));
        }
      } else {
        const r = stripImageExif(Buffer.from(data), name);
        strippedExifBytes += r.strippedBytes;
        if (r.strippedBytes > 0) touched.add(name);
        out.set(name, r.buf);
      }
    } else {
      // M4: word/embeddings/*.xlsm|*.xlsb|*.bin never reach this passthrough
      // under the default disposition — the macro gate above refuses them
      // (E_VALIDATION) unless --allow-macros preserves them verbatim or
      // --strip-macros removes them. Only macro-free parts flow here.
      out.set(name, Buffer.from(data));
    }
  }

  return {
    files: out,
    report: {
      ok: true,
      zeroRetention: o.zeroRetention,
      removedRSIDs,
      purgedAuthors,
      strippedExifBytes,
      purgedPII: propsPII + bodyPII,
      bodyPII,
      propsPII,
      acceptedRevisions,
      strippedComments,
      filesTouched: [...touched].sort(),
      droppedRogue: [...droppedRogue].sort(),
      scannedParts: [...scanned].sort(),
      scannedCount: scanned.size,
      macroParts: [...macroParts].sort(),
      strippedMacros: [...strippedMacros].sort(),
      strippedSvgActive: [...strippedSvg].sort(),
    },
  };
}

/**
 * Export entry point: sanitizes a DOCX buffer and returns a fresh buffer plus
 * a zero-retention-safe report (counts only — never original values).
 */
export function sanitizeDocxBuffer(
  docx: Buffer,
  opts?: SanitizeOptions
): { buffer: Buffer; report: SanitizeReport } {
  const o = resolveOptions(opts);
  const zip = new AdmZip(docx);
  checkZipBombCeilings(zip, { sourceCompressedBytes: docx.length });

  const files = new Map<string, Buffer>();
  const preDropped: string[] = [];
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) continue;
    if (isDroppedEntry(entry.entryName)) {
      if (!preDropped.includes(entry.entryName)) preDropped.push(entry.entryName);
      continue;
    }
    const data = Buffer.from(entry.getData());
    if (isXmlPart(entry.entryName)) {
      assertNoForbiddenMarkup(data.toString('utf8'), entry.entryName);
    }
    files.set(entry.entryName, data);
  }
  const { files: clean, report } = sanitizePackageEntries(files, o);
  for (const name of preDropped) {
    if (!report.droppedRogue.includes(name)) report.droppedRogue.push(name);
    if (!report.filesTouched.includes(name)) report.filesTouched.push(name);
  }
  report.droppedRogue.sort();
  report.filesTouched.sort();
  const next = new AdmZip();
  for (const [name, data] of clean) next.addFile(name, data);
  clampZipEntryTimes(next);
  return { buffer: next.toBuffer(), report };
}

/**
 * Non-throwing macro scan over a raw DOCX buffer (zip names + rel contents).
 * Used by read-only gates (e.g. PDF export) that never build a parser.
 */
export function listMacroPartsInBuffer(docx: Buffer): string[] {
  const zip = new AdmZip(docx);
  const names: string[] = [];
  for (const entry of zip.getEntries()) {
    if (!entry.isDirectory) names.push(entry.entryName);
  }
  return findMacroPartNames(names, (n) => {
    try {
      const e = zip.getEntry(n);
      return e ? e.getData().toString('utf8') : null;
    } catch {
      return null;
    }
  });
}

/**
 * Strips macro parts from a raw DOCX buffer (explicit-strip entry point for
 * flows that never build a parser). Call only after an explicit strip choice;
 * returns the re-packed buffer plus removed/edited entry names.
 */
export function stripMacroPartsFromBuffer(docx: Buffer): {
  buffer: Buffer;
  removed: string[];
  edited: string[];
} {
  const zip = new AdmZip(docx);
  checkZipBombCeilings(zip, { sourceCompressedBytes: docx.length });
  const files = new Map<string, Buffer>();
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) continue;
    files.set(entry.entryName, Buffer.from(entry.getData()));
  }
  const { removed, edited } = applyMacroStrip(mapMacroStore(files));
  const next = new AdmZip();
  for (const [name, data] of files) next.addFile(name, data);
  clampZipEntryTimes(next);
  return { buffer: next.toBuffer(), removed, edited };
}
