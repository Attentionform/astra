import { AdmZip, assertValidDocx, clampZipEntryTimes, SANITIZER_DEFAULT_AUTHOR, redactAttributePii, redactBodyPiiInWordText, redactTextNodePii, sanitizeAppPropertiesXml, sanitizeCorePropertiesXml, sanitizeCustomPropertiesXml, sanitizeDocxBuffer, validateSanitizeOptionsStrict } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { macroFlagsUsageError, readDocx, writeDocx } from './io';

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
    let comments = 0;
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
      const refs = xml.match(/<w:comment(RangeStart|RangeEnd|Reference)[\s>/]/g);
      if (refs) comments += refs.length;
      if (name === 'word/comments.xml' || name === 'word/commentsExtended.xml') {
        const defs = xml.match(/<w:comment[\s>/]/g);
        if (defs) comments += defs.length;
      }
    }
    return { revisions, comments };
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
  let allowMacros = false;
  let stripMacros = false;
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
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a.startsWith('--')) return usage(out, `sanitize: unknown flag "${a}"`);
    else if (!input) input = a;
    else if (!output) output = a;
    else return usage(out, 'sanitize takes exactly <in.docx> <out.docx>');
  }
  if (!input || !output) return usage(out, 'sanitize requires <in.docx> <out.docx>');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  if (replacementAuthor !== undefined && replacementAuthor.trim() === '') {
    return usage(out, 'sanitize --replacement-author must be non-empty');
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
  const effAcceptAll = acceptAllRevisions || preset;
  const effStripComments = stripComments || preset;
  const effAuthor = replacementAuthor ?? (preset ? SANITIZER_DEFAULT_AUTHOR : undefined);
  const doBodyRedact = (clientSafe || redactPii) && !metadataOnly;
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
      await assertValidDocx(buffer, 'Sanitized document');
      await writeDocx(output, buffer);
      const bodyLocations = collectBodyPiiLocations(buf);
      const locDisplay =
        bodyLocations.length > 0
          ? bodyLocations.slice(0, 20).join(',') + (bodyLocations.length > 20 ? `,+${bodyLocations.length - 20} more` : '')
          : 'none';
      print(
        out,
        `SANITIZED ${input} -> ${output} ` +
          `(rsids=${report.removedRSIDs} authors=${report.purgedAuthors} ` +
          `exifBytes=${report.strippedExifBytes} pii=${report.purgedPII} ` +
          `bodyRedactions=${report.bodyPII} bodyLocations=${locDisplay} ` +
          `revisions=${report.acceptedRevisions} comments=${report.strippedComments} files=${report.filesTouched.length})`
      );
      // Partial flags leave content behind by request — say so loudly on stderr
      // (fail-closed DX: the caller must see exactly what survives). Warn only
      // when markup actually survived in the output (counts > 0); a clean file
      // stays quiet instead of a flag-semantics warning.
      const survivors: string[] = [];
      if (!effAcceptAll || !effStripComments) {
        const lived = countSanitizeSurvivors(buffer);
        if (!effAcceptAll && lived.revisions > 0)
          survivors.push(`tracked revisions (${lived.revisions}) survive (w:ins/w:del)`);
        if (!effStripComments && lived.comments > 0) survivors.push(`comments (${lived.comments}) survive`);
      }
      if (survivors.length > 0 && !out.json) {
        printStderr(
          `${CLI_NAME}: sanitize WARNING: ${survivors.join(' + ')}. Remedy: re-run with --client-safe (or --accept-all-revisions --strip-comments) for a fully clean file`
        );
      }
      printJson(out, {
        ok: true,
        input,
        output,
        report,
        bodyLocations,
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
      report.propsPII = extraPropsPII;
      report.purgedPII = extraPropsPII + report.bodyPII;
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
    print(
      out,
      `SANITIZED ${input} -> ${output} ` +
        `(rsids=${report.removedRSIDs} authors=${report.purgedAuthors} ` +
        `exifBytes=${report.strippedExifBytes} pii=${report.purgedPII} ` +
        `bodyRedactions=${report.bodyPII} bodyLocations=none metadataOnly ` +
        `revisions=${report.acceptedRevisions} comments=${report.strippedComments} files=${report.filesTouched.length})`
    );
    const survivors: string[] = [];
    {
      const lived = countSanitizeSurvivors(finalBuffer);
      if (!effAcceptAll && lived.revisions > 0)
        survivors.push(`tracked revisions (${lived.revisions}) survive (w:ins/w:del)`);
      if (!effStripComments && lived.comments > 0) survivors.push(`comments (${lived.comments}) survive`);
    }
    survivors.push('body PII survives (body text kept; re-run without --metadata-only with --redact-pii or --client-safe to redact)');
    if (!out.json) {
      printStderr(
        `${CLI_NAME}: sanitize WARNING: ${survivors.join(' + ')}. Remedy: re-run without --metadata-only (with --redact-pii or --client-safe) for full body redaction`
      );
    }
    printJson(out, {
      ok: true,
      input,
      output,
      report,
      bodyLocations,
      metadataOnly: true,
      warnings: survivors,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file: input, fallbackCode: 'E_VALIDATION' });
  }
}
