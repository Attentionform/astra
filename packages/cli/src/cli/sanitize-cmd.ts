import * as fsp from 'fs/promises';
import * as path from 'path';
import { AdmZip, assertValidDocx, clampZipEntryTimes, SANITIZER_DEFAULT_AUTHOR, PII_COVERAGE, piiCoverageSentence, redactAttributePii, redactBodyPiiInWordText, redactTextNodePii, sanitizeAppPropertiesXml, sanitizeCorePropertiesXml, sanitizeCustomPropertiesXml, sanitizeDocxBuffer, validateSanitizeOptionsStrict } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { macroFlagsUsageError, macroOutputRefusal, readDocx, writeDocx } from './io';

/**
 * The note that follows a sanitize summary. `pii=0 bodyRedactions=0` alone reads as "no personal data", but it
 * only means that none of the patterns the check knows was found, so the note names them, and says what the
 * two counts count.
 */
function piiNote(report: { purgedPII: number; bodyPII: number }, bodyChecked: boolean): string {
  const counts = bodyChecked
    ? `pii=${report.purgedPII} counts document properties and bodyRedactions=${report.bodyPII} counts body text.`
    : `pii=${report.purgedPII} counts document properties only; the body text was not checked.`;
  return `${CLI_NAME}: sanitize NOTE: ${piiCoverageSentence()} ${counts}`;
}

// ── R-9 sanitize disclosure + metadata-only helpers (sanitize path only) ───
// Bare sanitize keeps body text (purgePII false for body). The SANITIZED
// stdout line discloses the count as bodyRedactions=N plus structural node
// locations (zip part + paragraph index, never original values).
// --metadata-only skips body-text redaction while metadata/rsids/authors/
// comments cleaning still applies; docProps PII is still purged as metadata.
// --metadata-only is an alias of bare; --redact-pii / --client-safe opt into body redaction.
export function isSanitizeBodyPartName(name: string): boolean {
  return (
    name === 'word/document.xml' || /^word\/(header|footer|footnotes|endnotes).*\.xml$/.test(name)
  );
}

export function isSanitizeCommentFamilyName(name: string): boolean {
  return (
    name === 'word/comments.xml' ||
    name === 'word/commentsExtended.xml' ||
    name === 'word/people.xml'
  );
}

export function isSanitizeStyleFamilyName(name: string): boolean {
  return name === 'word/styles.xml' || name === 'word/numbering.xml';
}

export function isSanitizeGenericTextName(name: string): boolean {
  return (
    /^word\/customXml\/.*\.xml$/.test(name) ||
    /^word\/charts\/.*\.xml$/.test(name) ||
    isSanitizeStyleFamilyName(name)
  );
}

/**
 * Count revision/comment markup surviving in a sanitized package (count-only,
 * never echoes values). Revisions cover w:ins/w:del/move + cell markers +
 * range markers + *Change wrappers; comments cover range/reference elements
 * plus comment definitions in comment parts. Used to gate the partial-flag
 * WARNING so a clean file stays quiet.
 */
export function countSanitizeSurvivors(docxBuf: Buffer): { revisions: number; comments: number } {
  try {
    const zip = new AdmZip(docxBuf);
    let revisions = 0;
    let commentDefinitions = 0;
    let commentReferences = 0;
    for (const entry of zip.getEntries()) {
      if (entry.isDirectory) continue;
      const name = entry.entryName;
      if (!name.startsWith('word/') || !name.endsWith('.xml')) continue;
      let xml: string;
      try {
        xml = entry.getData().toString('utf8');
      } catch {
        continue;
      }
      if (!xml.includes('<w:')) continue;
      const core = xml.match(
        /<w:(ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd|customXmlInsRangeStart|customXmlInsRangeEnd|customXmlDelRangeStart|customXmlDelRangeEnd|customXmlMoveFromRangeStart|customXmlMoveFromRangeEnd|customXmlMoveToRangeStart|customXmlMoveToRangeEnd)[\s>/]/g
      );
      if (core) revisions += core.length;
      const changes = xml.match(/<w:[A-Za-z0-9]+Change[\s>/]/g);
      if (changes) revisions += changes.length;
      // A comment is one definition in comments.xml and one reference in the text. Counting its range markers too
      // reported four for each comment.
      const refs = xml.match(/<w:commentReference[\s>/]/g);
      if (refs) commentReferences += refs.length;
      if (name === 'word/comments.xml') {
        const defs = xml.match(/<w:comment[\s>/]/g);
        if (defs) commentDefinitions += defs.length;
      }
    }
    return { revisions, comments: Math.max(commentDefinitions, commentReferences) };
  } catch {
    return { revisions: 0, comments: 0 };
  }
}

/**
 * Structural locations of body PII that sanitize will redact (count-only,
 * never echoes values). One entry per redaction hit (`part#p<idx>` for
 * body/comment parts, `part#text` for generic text carriers). Bounded at
 * 1000 entries. Falls back to part-level markers when paragraph slicing
 * finds nothing but the whole-part count is non-zero, so a non-zero
 * bodyRedactions count never prints with an empty location list.
 */
export function collectBodyPiiLocations(inputBuf: Buffer): string[] {
  const locations: string[] = [];
  let zip: InstanceType<typeof AdmZip>;
  try {
    zip = new AdmZip(inputBuf);
  } catch {
    return [];
  }
  const pushCapped = (loc: string): void => {
    if (locations.length < 1000) locations.push(loc);
  };
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) continue;
    const name = entry.entryName;
    const isBody = isSanitizeBodyPartName(name) || isSanitizeCommentFamilyName(name);
    const isGeneric = isSanitizeGenericTextName(name);
    if (!isBody && !isGeneric) continue;
    let xml: string;
    try {
      xml = entry.getData().toString('utf8');
    } catch {
      continue;
    }
    if (!xml.includes('<')) continue;
    if (isBody) {
      const paraRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
      const re = new RegExp(paraRe.source, 'g');
      let idx = 0;
      let m: RegExpExecArray | null;
      let paraHits = 0;
      while ((m = re.exec(xml)) !== null) {
        const paraXml = m[0];
        let count = 0;
        try {
          count += redactBodyPiiInWordText(paraXml, false).count;
          count += redactTextNodePii(paraXml, false).count;
          if (isSanitizeCommentFamilyName(name)) count += redactAttributePii(paraXml, false).count;
        } catch {
          count = 0;
        }
        for (let k = 0; k < count; k++) pushCapped(`${name}#p${idx}`);
        paraHits += count;
        idx++;
        if (locations.length >= 1000) break;
      }
      if (idx === 0 || paraHits === 0) {
        try {
          const whole =
            redactBodyPiiInWordText(xml, false).count +
            redactTextNodePii(xml, false).count +
            (isSanitizeCommentFamilyName(name) ? redactAttributePii(xml, false).count : 0);
          const extra = whole - paraHits;
          for (let k = 0; k < extra; k++) pushCapped(`${name}#outside-p`);
        } catch {
          /* count-only best effort */
        }
      }
    } else {
      try {
        let whole = redactTextNodePii(xml, false).count;
        if (isSanitizeStyleFamilyName(name)) whole += redactAttributePii(xml, false).count;
        for (let k = 0; k < whole; k++) pushCapped(`${name}#text`);
      } catch {
        /* count-only best effort */
      }
    }
    if (locations.length >= 1000) break;
  }
  return locations;
}

/** True when both paths name one file: the same path, a link to it, or a second name (hard link) for it. */
async function isSameFile(a: string, b: string): Promise<boolean> {
  if (path.resolve(a) === path.resolve(b)) return true;
  try {
    const [sa, sb] = await Promise.all([fsp.stat(a), fsp.stat(b)]);
    return sa.dev === sb.dev && sa.ino === sb.ino;
  } catch {
    return false;
  }
}

async function pathExists(p: string): Promise<boolean> {
  try {
    await fsp.lstat(p);
    return true;
  } catch {
    return false;
  }
}

/** What the run did with review history, and which of the choices the user made by name. */
interface ReviewChoices {
  /** The run left tracked changes in the file (it did not accept them). */
  keptRevisions: boolean;
  /** The run left comments in the file (it did not remove them). */
  keptComments: boolean;
  /** `--keep-revisions` was given. */
  askedKeepRevisions: boolean;
  /** `--keep-comments` was given. */
  askedKeepComments: boolean;
}

/**
 * The SANITIZED line. `metadataOnly` is named only when `--metadata-only` was given: every run that does not
 * redact the body used to print it, which said "metadata only" about runs that never used the flag.
 */
function sanitizedLine(
  input: string,
  output: string,
  report: { removedRSIDs: number; purgedAuthors: number; strippedExifBytes: number; purgedPII: number; bodyPII: number; acceptedRevisions: number; strippedComments: number; filesTouched: string[] },
  bodyLocations: string,
  metadataOnly: boolean
): string {
  return (
    `SANITIZED ${input} -> ${output} ` +
    `(rsids=${report.removedRSIDs} authors=${report.purgedAuthors} ` +
    `exifBytes=${report.strippedExifBytes} pii=${report.purgedPII} ` +
    `bodyRedactions=${report.bodyPII} bodyLocations=${bodyLocations}${metadataOnly ? ' metadataOnly' : ''} ` +
    `revisions=${report.acceptedRevisions} comments=${report.strippedComments} files=${report.filesTouched.length})`
  );
}

/**
 * Sentences about the revisions and comments that are still in the output. A run that kept them because the user
 * asked says so and says what that means (deleted text is readable, dates stay); any other run that kept them
 * (an individual flag) says that they survive. Quiet when nothing survived.
 */
function reviewSurvivors(lived: { revisions: number; comments: number }, choices: ReviewChoices): string[] {
  const out: string[] = [];
  if (choices.keptRevisions && lived.revisions > 0) {
    out.push(
      choices.askedKeepRevisions
        ? `tracked revisions (${lived.revisions}) kept because of --keep-revisions; deleted text is still readable in this file and the revision dates are kept`
        : `tracked revisions (${lived.revisions}) survive (w:ins/w:del)`
    );
  }
  if (choices.keptComments && lived.comments > 0) {
    out.push(
      choices.askedKeepComments
        ? `comments (${lived.comments}) kept because of --keep-comments; the comment dates are kept`
        : `comments (${lived.comments}) survive`
    );
  }
  return out;
}

/** The flags that asked to keep review history, for a remedy that says what to remove. */
function keepFlagNames(choices: ReviewChoices): string[] {
  return [choices.askedKeepRevisions ? '--keep-revisions' : '', choices.askedKeepComments ? '--keep-comments' : ''].filter(Boolean);
}

export async function cmdSanitize(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'sanitize');
  let input: string | undefined;
  let output: string | undefined;
  let zeroRetention = false;
  let replacementAuthor: string | undefined;
  let acceptAllRevisions = false;
  let stripComments = false;
  let clientSafe = false;
  let redactPii = false;
  let metadataOnly = false;
  let keepRevisions = false;
  let keepComments = false;
  let allowMacros = false;
  let stripMacros = false;
  let force = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--zero-retention') zeroRetention = true;
    else if (a === '--client-safe') clientSafe = true;
    else if (a === '--redact-pii') redactPii = true;
    else if (a === '--replacement-author') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'sanitize --replacement-author requires a non-empty value');
      }
      replacementAuthor = args[++i];
    } else if (a.startsWith('--replacement-author=')) {
      const v = a.slice('--replacement-author='.length);
      if (v.trim() === '') {
        return usage(out, 'sanitize --replacement-author must be non-empty');
      }
      replacementAuthor = v;
    } else if (a === '--accept-all-revisions') acceptAllRevisions = true;
    else if (a === '--strip-comments') stripComments = true;
    else if (a === '--metadata-only') metadataOnly = true;
    else if (a === '--keep-revisions') keepRevisions = true;
    else if (a === '--keep-comments') keepComments = true;
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a === '--force') force = true;
    else if (a.startsWith('--')) return usage(out, `sanitize: unknown flag "${a}"`);
    else if (!input) input = a;
    else if (!output) output = a;
    else return usage(out, 'sanitize takes exactly <in.docx> <out.docx>');
  }
  if (!input || !output) return usage(out, 'sanitize requires <in.docx> <out.docx>');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  if (keepRevisions && acceptAllRevisions) {
    return usage(out, 'sanitize: --keep-revisions and --accept-all-revisions contradict each other; use one of them');
  }
  if (keepComments && stripComments) {
    return usage(out, 'sanitize: --keep-comments and --strip-comments contradict each other; use one of them');
  }
  if (replacementAuthor !== undefined && replacementAuthor.trim() === '') {
    return usage(out, 'sanitize --replacement-author must be non-empty');
  }
  // sanitize writes a new file: never over its own input, and over another file only with --force.
  if (await isSameFile(input, output)) {
    return usage(out, `sanitize: <in.docx> and <out.docx> are the same file ("${path.basename(output)}"); give a different output path`);
  }
  if (!force && (await pathExists(output))) {
    return usage(out, `sanitize: destination "${path.basename(output)}" already exists (refusing to overwrite without --force)`);
  }
  // Bare `sanitize in out` keeps 100% body text while still stripping
  // metadata/rsids/authors/comments/revisions (zero-retention +
  // accept-all-revisions + strip-comments + neutral author).
  // Body redaction needs explicit --redact-pii or --client-safe bundle.
  // Lone --redact-pii carries the same metadata preset (fully clean plus
  // body redaction); --metadata-only is an alias of bare and wins over
  // redaction flags for body text; docProps PII is still purged as metadata.
  const bare = !zeroRetention && !acceptAllRevisions && !stripComments && replacementAuthor === undefined && !clientSafe && !redactPii;
  const preset = bare || clientSafe || metadataOnly || redactPii;
  const effZeroRetention = zeroRetention || preset;
  // The keep flags are not option flags (a run with only --keep-revisions is still the metadata preset), they turn
  // one step of the preset off.
  const effAcceptAll = acceptAllRevisions || (preset && !keepRevisions);
  const effStripComments = stripComments || (preset && !keepComments);
  const effAuthor = replacementAuthor ?? (preset ? SANITIZER_DEFAULT_AUTHOR : undefined);
  const doBodyRedact = (clientSafe || redactPii) && !metadataOnly;
  const choices: ReviewChoices = {
    keptRevisions: !effAcceptAll,
    keptComments: !effStripComments,
    askedKeepRevisions: keepRevisions,
    askedKeepComments: keepComments,
  };
  let buf: Buffer;
  try {
    buf = await readDocx(input);
  } catch (err: unknown) {
    return failHuman(out, err, { file: input, fallbackCode: 'E_IO' });
  }
  try {
    if (doBodyRedact) {
      const opts = validateSanitizeOptionsStrict({
        ...(effZeroRetention ? { zeroRetention: true } : {}),
        ...(effAuthor !== undefined ? { replacementAuthor: effAuthor } : {}),
        ...(effAcceptAll ? { acceptAllRevisions: true } : {}),
        ...(effStripComments ? { stripComments: true } : {}),
        ...(allowMacros ? { allowMacros: true } : {}),
        ...(stripMacros ? { stripMacros: true } : {}),
      }).options;
      const { buffer, report } = sanitizeDocxBuffer(buf, opts);
      const refusal = macroOutputRefusal('sanitize', output, report.macroParts, stripMacros);
      if (refusal) return usage(out, refusal);
      await assertValidDocx(buffer, 'Sanitized document');
      await writeDocx(output, buffer);
      const bodyLocations = collectBodyPiiLocations(buf);
      const locDisplay =
        bodyLocations.length > 0
          ? bodyLocations.slice(0, 20).join(',') + (bodyLocations.length > 20 ? `,+${bodyLocations.length - 20} more` : '')
          : 'none';
      print(out, sanitizedLine(input, output, report, locDisplay, false));
      // Partial flags leave content behind by request — say so loudly on stderr
      // (fail-closed DX: the caller must see exactly what survives). Warn only
      // when markup actually survived in the output (counts > 0); a clean file
      // stays quiet instead of a flag-semantics warning.
      const survivors: string[] = [];
      if (!effAcceptAll || !effStripComments) {
        survivors.push(...reviewSurvivors(countSanitizeSurvivors(buffer), choices));
      }
      if (survivors.length > 0 && !out.json) {
        const asked = keepFlagNames(choices);
        printStderr(
          `${CLI_NAME}: sanitize WARNING: ${survivors.join(' + ')}. Remedy: ` +
            (asked.length > 0
              ? `re-run without ${asked.join(' and ')} for a file without review history`
              : 're-run with --client-safe (or --accept-all-revisions --strip-comments) for a fully clean file')
        );
      }
      if (!out.json) printStderr(piiNote(report, true));
      printJson(out, {
        ok: true,
        input,
        output,
        report,
        bodyLocations,
        keptRevisions: choices.keptRevisions,
        keptComments: choices.keptComments,
        piiCoverage: PII_COVERAGE,
        ...(survivors.length > 0 ? { warnings: survivors } : {}),
      });
      return EXIT_OK;
    }
    const optsNoBody = validateSanitizeOptionsStrict({
      ...(effZeroRetention ? { zeroRetention: true } : {}),
      ...(effAuthor !== undefined ? { replacementAuthor: effAuthor } : {}),
      ...(effAcceptAll ? { acceptAllRevisions: true } : {}),
      ...(effStripComments ? { stripComments: true } : {}),
      ...(allowMacros ? { allowMacros: true } : {}),
      ...(stripMacros ? { stripMacros: true } : {}),
      purgePII: false,
    }).options;
    const { buffer: intermediate, report } = sanitizeDocxBuffer(buf, optsNoBody);
    const refusal = macroOutputRefusal('sanitize', output, report.macroParts, stripMacros);
    if (refusal) return usage(out, refusal);
    let finalBuffer = intermediate;
    try {
      const zip = new AdmZip(intermediate);
      const propsOpts: { zeroRetention: boolean; replacementAuthor?: string; purgePII: boolean } = {
        zeroRetention: effZeroRetention,
        ...(effAuthor !== undefined ? { replacementAuthor: effAuthor } : {}),
        purgePII: true,
      };
      let extraPropsPII = 0;
      const wanted = ['docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'];
      let propsChanged = false;
      for (const name of wanted) {
        const entry = zip.getEntry(name);
        if (!entry) continue;
        const xml = entry.getData().toString('utf8');
        let nextXml = xml;
        let pii = 0;
        if (name === 'docProps/core.xml') {
          const r = sanitizeCorePropertiesXml(xml, propsOpts);
          nextXml = r.xml;
          pii = r.purgedPII;
        } else if (name === 'docProps/app.xml') {
          const r = sanitizeAppPropertiesXml(xml, propsOpts);
          nextXml = r.xml;
          pii = r.purgedPII;
        } else {
          const r = sanitizeCustomPropertiesXml(xml, propsOpts);
          nextXml = r.xml;
          pii = r.purgedPII;
        }
        extraPropsPII += pii;
        if (nextXml !== xml) {
          zip.updateFile(name, Buffer.from(nextXml, 'utf8'));
          propsChanged = true;
          if (!report.filesTouched.includes(name)) report.filesTouched.push(name);
        }
      }
      // This path runs the engine with purgePII false, so a people.xml that survives (comments were kept) would keep
      // the account id of every commenter (w15:userId, for example an email or a Windows SID) although the author
      // name is emptied. Redact the attributes of that one part, as the engine does when purgePII is on.
      let peoplePII = 0;
      const peopleEntry = zip.getEntry('word/people.xml');
      if (peopleEntry) {
        const peopleXml = peopleEntry.getData().toString('utf8');
        const r = redactAttributePii(peopleXml, true);
        peoplePII = r.count;
        if (r.xml !== peopleXml) {
          zip.updateFile('word/people.xml', Buffer.from(r.xml, 'utf8'));
          propsChanged = true;
          if (!report.filesTouched.includes('word/people.xml')) report.filesTouched.push('word/people.xml');
        }
      }
      report.propsPII = extraPropsPII + peoplePII;
      report.purgedPII = extraPropsPII + peoplePII + report.bodyPII;
      report.filesTouched.sort();
      if (propsChanged) {
        try {
          clampZipEntryTimes(zip);
        } catch {
          /* deterministic stamp best effort */
        }
        finalBuffer = zip.toBuffer();
      }
    } catch (err: unknown) {
      return failHuman(out, err, { file: input, fallbackCode: 'E_VALIDATION' });
    }
    await assertValidDocx(finalBuffer, 'Sanitized document');
    await writeDocx(output, finalBuffer);
    const bodyLocations: string[] = [];
    print(out, sanitizedLine(input, output, report, 'none', metadataOnly));
    const survivors: string[] = reviewSurvivors(countSanitizeSurvivors(finalBuffer), choices);
    // `--metadata-only` is the same as bare sanitize and names no review history. Say so when it just removed some.
    const hint =
      metadataOnly && !keepRevisions && !keepComments && (report.acceptedRevisions > 0 || report.strippedComments > 0)
        ? `${CLI_NAME}: sanitize NOTE: --metadata-only accepted ${report.acceptedRevisions} tracked changes and removed ${report.strippedComments} comments, as bare sanitize does. To keep them, add --keep-revisions and/or --keep-comments.`
        : null;
    survivors.push(
      metadataOnly
        ? 'body PII survives (body text kept; re-run without --metadata-only with --redact-pii or --client-safe to redact)'
        : 'body PII survives (body text kept; re-run with --redact-pii or --client-safe to redact)'
    );
    if (!out.json) {
      printStderr(piiNote(report, false));
      if (hint) printStderr(hint);
      const asked = keepFlagNames(choices);
      printStderr(
        `${CLI_NAME}: sanitize WARNING: ${survivors.join(' + ')}. Remedy: ` +
          (metadataOnly
            ? 're-run without --metadata-only (with --redact-pii or --client-safe) for full body redaction'
            : 're-run with --redact-pii or --client-safe for full body redaction') +
          (asked.length > 0 ? `; re-run without ${asked.join(' and ')} for a file without review history` : '')
      );
    }
    printJson(out, {
      ok: true,
      input,
      output,
      report,
      bodyLocations,
      ...(metadataOnly ? { metadataOnly: true } : {}),
      keptRevisions: choices.keptRevisions,
      keptComments: choices.keptComments,
      piiCoverage: PII_COVERAGE,
      warnings: survivors,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file: input, fallbackCode: 'E_VALIDATION' });
  }
}
