import { DocumentAST } from '../ast';

/**
 * BX-05/OX-12: allocates the next free bookmark w:id by scanning the live
 * document for the maximum in-use integer id (w:bookmarkStart/w:bookmarkEnd)
 * and returning max(scannedMax+1, highWater+1, 1). Replaces hardcoded w:id="0"
 * / fixed-offset bypasses so every bookmark in the package is unique.
 * Exported for reuse by the TOC manipulator and backend/operations/hyperlink.ts
 * (adopted D1R2).
 *
 * OX-12 follow-up (D1 brutal review): a pure scan-then-allocate still races
 * when two callers scan BEFORE either inserts (both observe the same max and
 * mint the same id — single-threaded sequential inserts are safe because the
 * second scan sees the first insert, but parallel/worker or pre-insert
 * batching is not). A process-wide high-water mark closes that window: the
 * returned id is max(scannedMax+1, highWater+1, 1) and highWater advances on
 * every allocation, so two pre-insert scans still mint distinct ids (the
 * second bumps past the first's reservation even though the DOM lacks it
 * yet). Over-allocation across documents is harmless (ids only grow).
 * The floor of 1 preserves the no-zero contract the hyperlink path relies
 * on (its retry stays effective), and the reservation applies in every
 * mode so same-input-twice stays stable via the reset hook in tests.
 */
let bookmarkHighWater = -1;

export function allocateNextBookmarkId(doc: Document): number {
  let maxId = -1;
  const tags = ['w:bookmarkStart', 'w:bookmarkEnd'];
  for (const tag of tags) {
    let els: ArrayLike<Element>;
    try {
      els = doc.getElementsByTagName(tag);
    } catch {
      continue;
    }
    for (let i = 0; i < els.length; i++) {
      const raw = (els[i] as Element).getAttribute('w:id');
      if (raw == null || raw === '') continue;
      const n = parseInt(raw, 10);
      if (Number.isFinite(n) && n > maxId) maxId = n;
    }
  }
  const candidate = Math.max(maxId + 1, bookmarkHighWater + 1, 1);
  // Reserve even the first allocation so a second pre-insert scan bumps past
  // it; the retry in the hyperlink path stays effective in every mode.
  bookmarkHighWater = Math.max(bookmarkHighWater, candidate);
  return candidate;
}

/** Test hook: resets the bookmark high-water reservation (isolates tests). */
export function resetBookmarkHighWaterForTests(): void {
  bookmarkHighWater = -1;
}

/**
 * OX-12 owning-batch helper (D1 follow-up, ADOPTED D1R2): atomic hyperlink rId
 * allocator. backend/operations/hyperlink.ts used to mint
 * `rId_link_${Date.now()}` — two same-ms hyperlinks collided (proven:
 * identical ids, duplicate Relationship Ids in document.xml.rels). It now
 * calls this helper, which mints canonical `rIdN` ids via the safe allocator
 * PLUS a process-wide high-water mark so even two manipulators seeded from
 * the same pre-write snapshot cannot collide: the returned number is
 * max(allocator candidate, globalHighWater+1, map-miss), and the high-water
 * advances on every call (reserving the id even before the caller writes it
 * to the map/zip). Callers must write the returned id to ast.relationships
 * BEFORE the zip (hyperlink.ts does). Exported for reuse + tests.
 */
let hyperlinkRIdHighWater = 0;
const mintedHyperlinkRIds = new Set<string>();

export function allocateSafeHyperlinkRId(ast: DocumentAST): string {
  const manip = new RelationshipsManipulator(ast);
  let candidate = manip.generateSafeRId();
  let num = parseInt(candidate.replace(/^rId/, ''), 10);
  if (!Number.isFinite(num) || num < 1) num = Object.keys(ast.relationships ?? {}).length + 1;
  // Global reservation: bump past any id minted earlier in this process even
  // when the caller's AST snapshot predates those writes (separate objects).
  num = Math.max(num, hyperlinkRIdHighWater + 1);
  let guard = 0;
  while (
    (Object.prototype.hasOwnProperty.call(ast.relationships, `rId${num}`) ||
      mintedHyperlinkRIds.has(`rId${num}`)) &&
    guard < 100000
  ) {
    num += 1;
    guard += 1;
  }
  candidate = `rId${num}`;
  hyperlinkRIdHighWater = Math.max(hyperlinkRIdHighWater, num);
  mintedHyperlinkRIds.add(candidate);
  return candidate;
}

/** Test hook: resets hyperlink allocator state (isolates tests). */
export function resetHyperlinkRIdStateForTests(): void {
  hyperlinkRIdHighWater = 0;
  mintedHyperlinkRIds.clear();
}

export class RelationshipsManipulator {
  private ast: DocumentAST;
  private maxRId: number = 0;

  constructor(ast: DocumentAST) {
    this.ast = ast;
    this.calculateMaxRId();
  }

  private calculateMaxRId() {
    // BX-05: seed from trailing digits of ANY rId shape (rId12, rIdH1 is
    // skipped — no trailing run after strip... see below) so bypass ids minted
    // outside this allocator still push the counter forward. Canonical
    // /^rId(\d+)$/ ids seed directly; non-canonical shapes (rId_link_<ts>,
    // rIdH1) seed from their trailing digit run when present.
    for (const rId of Object.keys(this.ast.relationships)) {
      const canonical = /^rId(\d+)$/.exec(rId);
      const trailing = /(\d+)$/.exec(rId);
      const m = canonical ?? trailing;
      if (!m) continue;
      const num = parseInt(m[1], 10);
      if (!isNaN(num) && num > this.maxRId) {
        this.maxRId = num;
      }
    }
  }

  /**
   * Recomputes the counter from the current AST map. Call after external
   * (zip-level) relationship edits so the allocator cannot reuse an id.
   */
  public reseed(): void {
    this.maxRId = 0;
    this.calculateMaxRId();
  }

  // Generates a new safe relationship ID without recycling old ones (prevents corruption)
  public generateSafeRId(): string {
    // BX-05: loop on existence — monotonic increment alone can still collide
    // with bypass ids (timestamp rId_link_*, comments-part counter) minted
    // outside this allocator. Never returns an id already in the AST map.
    // Guard: timestamp-seeded counters (Date.now() ~1.7e12) would mint absurd
    // rId1700000000001-style ids, so cap the trailing-digit seed influence:
    // when maxRId is implausibly large, fall back to map-size-based search.
    if (!Number.isFinite(this.maxRId) || this.maxRId > 100_000_000) {
      this.maxRId = Object.keys(this.ast.relationships).length;
    }
    let candidate = `rId${this.maxRId + 1}`;
    while (Object.prototype.hasOwnProperty.call(this.ast.relationships, candidate)) {
      this.maxRId++;
      candidate = `rId${this.maxRId + 1}`;
    }
    this.maxRId++;
    return candidate;
  }

  public addRelationship(type: string, target: string, targetMode: 'External' | 'Internal' = 'Internal'): string {
    const rId = this.generateSafeRId();
    this.ast.relationships[rId] = { type, target, targetMode };
    return rId;
  }

  public removeRelationship(rId: string): void {
    delete this.ast.relationships[rId];
  }

  public addMedia(fileName: string, buffer: Buffer): string {
    this.ast.mediaFiles[fileName] = buffer;
    return this.addRelationship(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      `media/${fileName}`
    );
  }
}
