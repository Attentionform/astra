import type { UnresolvedRevision } from './ooxml/revisions';

/** Parts are canonicalised before the sanitizer reads them, so tracked-change elements are always written with `w:`. */
const REVISION_ELEMENT_RE = /<w:(ins|del|moveFrom|moveTo)(?=[\s/>])((?:[^<>"']|"[^"]*"|'[^']*')*?)>/g;
const REVISION_ID_RE = /\sw:id\s*=\s*(?:"([^"]*)"|'([^']*)')/;
/** Comments, CDATA sections and processing instructions can hold text that looks like a tag. */
const OPAQUE_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g;

/**
 * After accept-all, the only tracked-change elements allowed in a part are the ones the report lists as unresolved.
 * Two levels (design DN3, 8.5). An element with a `w:id` is allowed only when an entry for the same part and kind
 * carries that `revisionId`; each entry allows one element. An element without a `w:id`, or when the entries carry
 * no ids, falls back to a bounded count for the part and kind. Limit of the count level: it cannot tell two
 * revisions of the same kind in one part apart, so a reported revision that vanished while an unreported one of the
 * same kind stayed still balances.
 */
export function assertOnlyReportedRevisions(part: string, source: string, unresolved: readonly UnresolvedRevision[]): void {
  const xml = source.replace(OPAQUE_RE, '');
  const withId = new Map<string, number>(); // `${kind}|${id}` -> entries that carry that id
  const idless = new Map<string, number>(); // kind -> entries without an id (count level)
  for (const u of unresolved) {
    if (u.part !== part) continue;
    if (u.revisionId !== undefined) withId.set(`${u.kind}|${u.revisionId}`, (withId.get(`${u.kind}|${u.revisionId}`) ?? 0) + u.count);
    else idless.set(u.kind, (idless.get(u.kind) ?? 0) + u.count);
  }
  const takeIdless = (kind: string): boolean => {
    const n = idless.get(kind) ?? 0;
    if (n <= 0) return false;
    idless.set(kind, n - 1);
    return true;
  };
  REVISION_ELEMENT_RE.lastIndex = 0;
  for (let m = REVISION_ELEMENT_RE.exec(xml); m !== null; m = REVISION_ELEMENT_RE.exec(xml)) {
    const kind = m[1];
    const idMatch = REVISION_ID_RE.exec(m[2] ?? '');
    const id = idMatch ? (idMatch[1] ?? idMatch[2]) : undefined;
    let allowed = false;
    if (id !== undefined) {
      const key = `${kind}|${id}`;
      if ((withId.get(key) ?? 0) > 0) {
        withId.set(key, (withId.get(key) as number) - 1);
        allowed = true;
      } else {
        allowed = takeIdless(kind); // the report gave no id for this one: count level
      }
    } else {
      // The element has no id: count level against everything reported for the kind.
      allowed = takeIdless(kind);
      if (!allowed) {
        for (const [key, n] of withId) {
          if (n > 0 && key.startsWith(`${kind}|`)) {
            withId.set(key, n - 1);
            allowed = true;
            break;
          }
        }
      }
    }
    if (!allowed) {
      throw new Error(`E_VALIDATION: sanitize could not finish "acceptAllRevisions": ${part} still has a tracked change (<w:${kind}${id !== undefined ? ` w:id="${id}"` : ''}>) that the report does not list. Nothing was written. Open the file in Word, save it again, then retry.`);
    }
  }
}

