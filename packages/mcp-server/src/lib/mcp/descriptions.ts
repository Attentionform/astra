// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 756-921).
// Zero logic edits.
import type { Lean6ToolName, GitHistoryToolName, ParityToolName } from './tool-names';
import { AUDIT_RETENTION_NOTE, CHANGED_VALUES_MAX, CHANGED_VALUE_TEXT_MAX } from './diff';

// ── Shared description blocks (ECMA essentials + workflow + examples) ────────
export const MCP_ECMA_ESSENTIALS =
  'ECMA-376 essentials: colors are hex without # (e.g. 4F46E5, never #4F46E5); ' +
  'font sizes are half-points in w:sz (e.g. 28 = 14pt, 22 = 11pt); ' +
  'spacing/margins/cellMar use twips dxa (twentieths of a point, e.g. 240 = 12pt, 180 = 9pt); ' +
  'borders use eighths-of-a-point sz (e.g. 32 = 4pt left accent).';

export const MCP_WORKFLOW =
  'Workflow: open_document -> inspect_document -> search_content -> execute_batch -> save_document -> close_document. ' +
  'Always inspect first, batch second, save third. ' +
  'Single edit = 1-item batch. Never call twice when once suffices.';

export const MCP_ENVELOPE_NOTE = '';

function withSharedDocs(body: string): string {
  return `${body} ${MCP_ECMA_ESSENTIALS} ${MCP_WORKFLOW} ${MCP_ENVELOPE_NOTE}`;
}

export const LEAN6_DESCRIPTIONS: Record<Lean6ToolName, string> = {
  open_document: withSharedDocs(
    'Open a .docx file into a session (Map sessionId->DocHandle, LRU 10 docs / 200MB / 30min idle evict). ' +
      'Returns {ok:true,sessionId,revision,summary,evicted} where evicted:string[] lists sessions dropped ' +
      'by the capacity cap to make room (re-open them if needed; later calls with an evicted id fail ' +
      'with E_SESSION plus a close_document remedy). ' +
      'Default summary is the full structural summary including the complete word/styles.xml dump ' +
      '(large templates return a large payload; no trim applied). ' +
      'Pass {"summaryOnly":true} for a palette-only summary ({palette:{availableStyles,brandColors,primaryFont}} ' +
      'parsed from word/styles.xml, no stylesXml dump). summaryOnly defaults to false (omitted/false keeps ' +
      'the full dump, unchanged, so existing clients never break). ' +
      'Example input {"path":"documents/report.docx"} -> {"ok":true,"sessionId":"sess_abc","revision":1,"evicted":[]}. ' +
      'Open first; keep the sessionId for every later call. ' +
      'Macro/active content refuses by default (E_VALIDATION naming the parts): pass {"allowMacros":true} to preserve or {"stripMacros":true} to strip (never both). ' +
      'Session ids are server-minted (sess_<uuid>, DEPRECATED to supply sessionId: client-supplied values are ' +
      'rejected with E_VALIDATION; deterministic ids are test-only via --allow-client-session-ids). ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions, and duplicate client-supplied ids are rejected.'
  ),
  inspect_document: withSharedDocs(
    'Inspect a session: standard TOON outline (fidelity compact|standard|full, default standard) plus style palette passthrough ' +
      '{availableStyles,brandColors,primaryFont} parsed from word/styles.xml, plus document core properties ' +
      '`metadata` {title,creator,subject,description,keywords,category,lastModifiedBy,revision,created,modified} ' +
      'read from docProps/core.xml (absent when the docx has no core part). ' +
      'Renders the in-memory session state (includes unsaved execute_batch edits); `dirty:true` means the ' +
      'session differs from the last saved file, `revision` is the in-memory revision. ' +
      'Node ids are verbatim: paragraphs carry 8-hex w14:paraId (e.g. 3F2A9C1B), tables t_001.., sections sec_root, sec_2, sec_3.. by position; use them as-is for targets/anchors. ' +
      'A repeat paragraph sharing one native paraId reports a `<paraId>_dup` session id instead (e.g. 3F2A9C1B_dup): saving refuses until the repeat is deleted and its text reinserted for a fresh id. ' +
      'Tables render one row per table with grid cells as rNcM coordinates (r1c1 value | r1c2 value, rows joined by ;); merged spans are marked (spanN); merge covers are tagged hidden-cover (not directly editable, address the restart cell). ' +
      'Cell edits: the replace, style, namedStyle, clearFormat, hyperlink, bookmark and comment actions accept target "<tableId>:r<row>c<col>" (1-based, as the grid prints it; add ":p<n>" to pick one paragraph of a multi-paragraph cell). Other actions, including table, take the table id. Every row and column number counts from 1 (cell addresses, rowIndex, colIndex, columnIndex, row, col, index, afterIndex, fromRow/toRow/fromCol/toCol), so r3c2 and {rowIndex:3,colIndex:2} are the same cell. ' +
      'Chart nodes render style chart|chart3d with title + categories + series in the text ([Chart column] Title | cats: Q1,Q2 | ARR: 10,25), read from the chart part. ' +
      'Flag legend: [TBL] table [IMG] image/drawing [EQ] equation [CITE] citation [CMT] comment [TRK] tracked [LNK] link [BKMK] bookmark [L#] list level (1 is the top level; list ops count the same way) [+ADDED]/[~MODIFIED] diff; standard tier adds [B] bold [I] italic [C:hex] color [SZ:half-points] size [RS:name] run style [J:just] [N:num] numbering [HDR]/[ZEBRA] table look [CH:type]/[RID:id] chart. ' +
      'Palette lists every style id verbatim with no trim (large templates return the full list). ' +
      '3-tier workflow for large docs: (1) macro-overview inspect (compact, first window) -> ' +
      '(2) search_content to find a target node id -> (3) local zoom via {aroundNode:"<id>",window:5} or ' +
      'page with {offset,limit}. Windowing opts: offset (min 0, default 0), limit (1-200, default 50), ' +
      'aroundNode + window (1-20, default 5) returning [node-w,node+w], maxTokens (200-8000, cuts on a clean ' +
      'row boundary, always reported via truncated/hasMore). aroundNode and offset/limit are mutually ' +
      'exclusive (use one or the other). Windowed renders preserve the requested tier; the TOON ' +
      'fidelity marker stays lossy by design (TOON is a lossy projection). Docs under ~2000 tokens return the full TOON ' +
      '(unchanged); larger docs default to the first ~50-node window ' +
      '{totalNodes,returnedNodes,offset,limit,hasMore,nextOffset,toonText} with a navigation remedy. ' +
      'Full fidelity requires sectionId; full-fidelity errors name the available section ids. ' +
      'Example {"sessionId":"sess_abc","fidelity":"standard"} -> TOON text + palette. ' +
      'Use the palette style names verbatim (e.g. Corporate Callout); never invent styles. ' +
      'When the session carries a style guide, the envelope also includes its compact styleGuide summary ' +
      '{name,headerFill,headerText,chartPalette,calloutTheme}. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions, and duplicate client-supplied ids are rejected.'
  ),
  search_content: withSharedDocs(
    'Search live paragraph/table w:t text; returns [{id,preview}] node hits for anchors/targets. ' +
      'A table hit also lists `cells` (up to 5 addresses such as t_001:r3c1) for the cells that hold the match; pass one as the target of replace/style/namedStyle/clearFormat/hyperlink/bookmark/comment. ' +
      'Chart titles/series are NOT matched by search (they live in the chart part, not w:t); discover them in inspect TOON chart rows, then anchor by the chart node id. ' +
      'Image alt text and footnote/endnote bodies are NOT matched either (search covers body w:t text only); ' +
      'find images via inspect [IMG] rows and footnotes via their anchor ids. Search extension to those parts is a later-program candidate, not this surface. ' +
      'Default is an exact literal substring (case-insensitive); pass {"regex":true} to interpret the ' +
      'query as a regular expression (bounded: pattern <=200 chars, ~500ms evaluation budget, case-insensitive). ' +
      'Caps: first 50 hits, 120-char previews. ' +
      'Literal queries containing regex metacharacters return literal results plus a hint that regex:true exists. ' +
      'Example {"sessionId":"sess_abc","query":"Executive Summary"} -> {"ok":true,"matches":[{"id":"3F2A9C1B","preview":"Executive ..."}]}. ' +
      'Search before editing when the anchor id is unknown.'
  ),
  execute_batch: withSharedDocs(
    'Apply 1..500 strictly-typed operations atomically (strict discriminatedUnion per action; sequential:true default preserves document order; ' +
      'single edit = 1-item batch). Caps 500 ops / 200MB throw E_QUOTA with remedy (maxSteps is an optional positive integer, unbounded when omitted). ' +
      'Client-supplied op ids are honored where the schema allows (insert/equation/footnote/chart accept `id`; strict schemas reject it elsewhere). ' +
      'Style ops support word-level spans: `match` styles only the matched substring (run splitting; `matchAll` styles ' +
      'every occurrence, `matchCase:false` matches case-insensitively); an absent match fails with E_NOT_FOUND. ' +
      'Returns {ok:true,appliedCount,revision,skippedIds,createdNodeIds,results:[{index,action,createdNodeIds}],revisionsPending,revisionsRemedy?} where createdNodeIds ' +
      'lists new node ids in document order (use them as anchors in the next batch without re-inspecting). ' +
      'A batch that drops a non-empty redo stack signals it (redoCleared:true + redoClearedWarning naming the depth): ' +
      'a fresh edit invalidates redo, so redo_document afterwards reports an empty stack. ' +
      'Example {"sessionId":"sess_abc","operations":[{"action":"replace","target":"p_001","text":"Hello"}],"sequential":true} -> ' +
      '{"ok":true,"appliedCount":1,"revision":2}. On failure returns {ok:false,code,remedy} with a fix-and-retry hint. ' +
      'Image inserts take exactly one source: {action:"image",operation:"insert",anchor,fileData} for inline base64 bytes, ' +
      'or {action:"image",operation:"insert",anchor,filePath} for a file read from a directory jail (both together fail with E_VALIDATION). ' +
      'The jail resolves explicit allowedDir first, else the session anchor (dirname of the open document); ' +
      'outside-jail refs fail with E_SANDBOX, missing files with E_IO, oversize files with E_LIMIT; batch reads nothing otherwise. ' +
      'Callout themes (operations/callout.ts THEMES): indigo|emerald|amber|slate|rose (default indigo; unknown names render indigo). ' +
      'Optional styleGuide (bundled guide name or guide .md path, jailed like filePath: explicit allowedDir else the session anchor; remote URLs fail closed) ' +
      'overrides the session guide for this batch only; it never replaces or persists (to re-skin the session use compose_document with styleGuide). ' +
      'Omitted inherits the session guide. Guides cannot be detached: an empty styleGuide fails with E_VALIDATION; open a new session to work without a guide. ' +
      'Unknown names fail with E_VALIDATION naming the available set. Guide parse notes ride both server stderr and the client-visible warnings array. ' +
      'Layout operations: {action:"section",operation:"setColumns",columns,spacing,separator} (columns 1-12, spacing in twips, separator draws the dividing rule), ' +
      '{action:"section",operation:"split"|"merge",...}, and {action:"pageSetup",orientation:"portrait"|"landscape",paperSize:"A4"|"Letter"|"Legal"|"A3",marginTop,marginBottom,marginLeft,marginRight,columns,sectionIndex} (margins in twips and written as given, never rotated; orientation turns the page by swapping its width and height; sectionIndex counts sections from 1 and limits the change to one section, omit it to change every section). ' +
      'Section breaks: {action:"insert",type:"sectionBreak",anchor,position,breakType:"nextPage"|"continuous"|"evenPage"|"oddPage"|"nextColumn"}; the section before the break keeps the page setup and headers of the section it splits. ' +
      'Headers and footers: {action:"headerFooter",operation:"set",targetType:"header"|"footer",text,alignment,sectionIndex (from 1, default the first section)}; {PAGE} and {NUMPAGES} inside text become live page-number fields. ' +
      'Images: width and height are integers in EMU (914400 per inch, 12700 per point, 9525 per pixel at 96 dpi); give one and the other follows the image aspect ratio, give neither and the image keeps its own size fitted to the page. ' +
      'Tables created by insert type:"table" or by Markdown repeat their header row across pages by default. ' +
      'Equations: {action:"equation",operation:"insert"|"replace",latex,mode:"auto"|"latex"|"plain"}; LaTeX that contains a backslash, ^ or _ becomes a real Word equation, and mode:"plain" keeps the text exactly as written. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions, and duplicate client-supplied ids are rejected.'
  ),
  save_document: withSharedDocs(
    'Validate then persist the session docx (validator-gated save: throws E_VALIDATION with xPath + table/node location + remedy when the package is invalid, never writes corruption). ' +
      'Save-blocked remedies are xpath-specific (drawing docPr/id, table tbl/tr/tc, relationships, header/footer parts, session tracking-attribute leaks each name their own fix). ' +
      'Optional commit params: message (commit subject, single-line, max 500 chars; newlines become spaces; ' +
      'blank/omitted falls back to `save_document: rev <n> (<k> ops since last save)`); author (max 100 chars; ' +
      'blank/omitted falls back to ASTra Agent). Over-long message/author fail at the handler with E_VALIDATION ' +
      'plus a shorten-and-retry remedy (never a bare MCP -32602: caps are enforced in the handler, not the wire schema). ' +
      'A successful save clears the session dirty flag (revision is re-synced to the file on disk). ' +
      'The success envelope also carries revisionsPending (w:ins+w:del left in the saved package) plus a revisionsRemedy naming accept when >0. ' +
      'The success envelope also carries advisory changedValues [{nodeId,before,after}] (old saved text -> new live text, ' +
      `bounded: at most ${CHANGED_VALUES_MAX} entries, each side at most ${CHANGED_VALUE_TEXT_MAX} chars, truncated flag honest; ` +
      'empty + reason when no baseline is computable). ' +
      'When the session carries a style guide (new_document/compose_document styleGuide), the envelope also carries ' +
      'advisory styleWarnings [{nodeId,rule,message}] (table contrast, chart duplicate colors, split-table orphans), never blocking. ' +
      'Example {"sessionId":"sess_abc"} -> {"ok":true,"path":"documents/report.docx"}. ' +
      'Always save after execute_batch; close afterwards. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions, and duplicate client-supplied ids are rejected.'
  ),
  close_document: withSharedDocs(
    'Evict a session and release its lock/memory. ' +
      'Returns {ok:true,sessionId,dirty,discardedRevisions,warning?}: a dirty close warns ' +
      '(dirty:true + warning naming the discarded unsaved revision count) instead of silently discarding; ' +
      'save first when the warning matters. ' +
      'Example {"sessionId":"sess_abc"} -> {"ok":true,"dirty":false,"discardedRevisions":0}. ' +
      'Close when finished; expired sessions (30min idle) must be re-opened.'
  ),
};

export const GIT_HISTORY_DESCRIPTIONS: Record<GitHistoryToolName, string> = {
  get_document_history: withSharedDocs(
    'List the saved git history for a session (newest first, read-only: never mutates, never flips dirty). ' +
      'Paginate with limit (1-100, default 20) and offset (default 0). ' +
      'Returns {ok:true,sessionId,documentPath,currentCommit,totalCommits,history:[{oid,message,author,timestamp,timestampMs,opCount,affectedNodeIds,resultingHeadOid}]} ' +
      '(timestampMs = ms-precision epoch millis; commit-time clock for tracked commits, git seconds for older ones; ' +
      'opCount is always present (0 for pre-tracking commits and heals-only saves: it folds operations, not normalizer rewrites); ' +
      'affectedNodeIds is always present ([] when untracked); resultingHeadOid always equals oid (append-only; refs never move). ' +
      'Example {"sessionId":"sess_abc","limit":10} -> history slice. ' +
      'Use the oids with revert_document (or preview_revert first) to restore an earlier saved state. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions, and duplicate client-supplied ids are rejected.'
  ),
  preview_revert: withSharedDocs(
    'Read-only WOULD-restore preview for revert_document (never mutates: no revision bump, no disk write, no redo-stack change). ' +
      'Target one commit: {steps:N} walks back N saved commits from the head (1 = the previous saved commit, must be < totalCommits) ' +
      'OR {commitOid:"<40-hex>"} jumps to that commit (mutually exclusive). ' +
      'Returns {ok:true,sessionId,targetOid,targetMessage,targetAuthor,stepsBack,affectedFiles,changedValues,changedValuesTruncated,changedValuesPositional,dirty,dirtyHandling,currentHeadOid,totalCommits} ' +
      'where affectedFiles lists package paths that differ from the target (word/document.xml plus tracked extras). ' +
      'The preview also carries advisory changedValues [{nodeId,before,after}] (live current text -> target restore text, ' +
      `bounded: at most ${CHANGED_VALUES_MAX} entries, each side at most ${CHANGED_VALUE_TEXT_MAX} chars, truncated flag honest; ` +
      'changedValuesPositional:true means either side held orphan ids (no shared id across the boundary, common under duplicate native ids) ' +
      'so before/after pairing is positional, not identity; ' +
      'empty + changedValuesReason when the target range is not computable). ' +
      'Example {"sessionId":"sess_abc","steps":1} -> preview summary. ' +
      'Preview first when a revert would discard unsaved edits.'
  ),
  revert_document: withSharedDocs(
    'Restore an earlier SAVED state as a new head commit (history is append-only; refs never move). ' +
      'Target one commit: {steps:N} walks back N saved commits from the head (1 = the previous saved commit, must be < totalCommits) ' +
      'OR {commitOid:"<40-hex>"} jumps to that commit (mutually exclusive). Bare reverts fail closed: unsaved execute_batch edits refuse with E_CONFLICT ' +
      'unless discardUnsaved:true, which discards them. ' +
      'The conflict remedy names net-zero edits (pending revisions count as dirty even when content matches) and carries dirtyReason with the pending count. ' +
      'Discarded unsaved edits are echoed back as {wasDirty,discardedNodeCount,discardedNodeIds} (counts + ids only, bounded), never silently dropped. ' +
      'Optional author (<=100 chars, default ASTra Agent) + reason (<=500 chars): when either is given the ' +
      'commit subject follows `Revert [N steps] to <oid> by <author>[: reason]`; otherwise the legacy ' +
      '`Revert to <oid>[: message]` shape is kept (message still accepted), and a bare default revert quotes ' +
      'the target commit message in the new subject (`Revert to <oid>: "<target message>"`). ' +
      'Example {"sessionId":"sess_abc","steps":1} -> {ok:true,revertedToCommit,revision,dirty:false,discardedUnsaved,commitOid,commitMessage}. ' +
      'The pre-revert state is kept for redo_document. ' +
      'Selective middle-undo recipe (no cherry-pick: restore the older keeper, then re-apply the kept newer edits): ' +
      '1) preview_revert the target, 2) revert_document to the older keeper commit, ' +
      '3) inspect/search for fresh anchors (node ids die on every revert; never reuse pre-revert ids), ' +
      '4) execute_batch to re-apply the kept newer edits, 5) save_document. ' +
      'Export-check after every mutation: run export_document (or inspect_document) after each revert and each re-apply ' +
      'and confirm the bytes before the next step.'
  ),
  redo_document: withSharedDocs(
    'Re-apply the state undone by the last revert_document (pops the per-session redo stack, restores it, commits as new head). ' +
      'Success carries advisory changedValues [{nodeId,before,after}] (pre-redo live text -> restored text, same bounds as save/preview). ' +
      'Empty stack fails with E_NOT_FOUND plus atLatestState:true (a fresh execute_batch edit clears redo, and each redo consumes one entry). ' +
      'Optional author (<=100 chars, default ASTra Agent) + reason (<=500 chars, preferred over message) with the same hygiene as save/revert. ' +
      'Example {"sessionId":"sess_abc"} -> {ok:true,revision,dirty:false}. ' +
      'A successful execute_batch clears the redo stack (fresh edits invalidate redo).'
  ),
};

/**
 * export_audit_trail description (handler-only: NOT a registered MCP tool, so
 * the 12-tool pin is unaffected). Carries the HOLD retention note verbatim.
 */
export const EXPORT_AUDIT_TRAIL_DESCRIPTION: string = withSharedDocs(
  'Export the full audit trail for a session: every saved commit (newest first) with author, ms-precision ' +
    'timestamps, op counts, affected node ids, and a sha256 packageHash over the committed document.xml per ' +
    'revision (each hash recomputed via read-only checkout, so verifiers can reproduce it). ' +
    'Returns {ok:true,sessionId,documentPath,totalCommits,currentCommit,trail:[{oid,message,author,timestampMs,opCount,affectedNodeIds,resultingHeadOid,packageHash}],retentionNote}. ' +
    AUDIT_RETENTION_NOTE
);

export const PARITY_DESCRIPTIONS: Record<ParityToolName, string> = {
  validate_document: withSharedDocs(
    'Preflight read-only OOXML audit for a session: runs the same validator gate save_document enforces, ' +
      'before anything is written. Detects schema violations while they are still cheap to fix. ' +
      'Strictly read-only: zero session mutation (revision and dirty flag untouched), zero disk writes, ' +
      'zero history commits. Returns {ok:true,sessionId,revision,dirty,validation:{valid,errorCount,errors}} ' +
      'where errors carry description + xPath (first 10, same shape as the save-blocked envelope). ' +
      'The verdict is on the document as a save would write it (children put in schema order first); when that save would repair something the result also carries pendingNormalizations (a list of repairs, absent when none), so a valid verdict is not mistaken for "valid exactly as it stands". ' +
      'Use before save_document when the package may be non-conforming; save_document re-gates before writing. ' +
      'Example {"sessionId":"sess_abc"} -> {"ok":true,"validation":{"valid":true,"errorCount":0,"errors":[]}}. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions.'
  ),
  get_stats: withSharedDocs(
    'Token-cheap document census for a session: words, paragraphs, tables, images, and a 500-words-per-page ' +
      'page estimate in a single small envelope. Counters are the exact engine counters behind CLI stats, ' +
      'so both surfaces always agree. Read-only: never mutates, never bumps the revision. ' +
      'Returns {ok:true,sessionId,revision,dirty,words,paragraphs,tables,images,pages}. ' +
      'Use instead of inspect_document when only counts are needed (a full TOON dump costs thousands of tokens ' +
      'for the same question). Example {"sessionId":"sess_abc"} -> {"ok":true,"words":42,"paragraphs":5}. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions.'
  ),
  sanitize_document: withSharedDocs(
    'Scrub an open session document: removes revision ids, authorship, embedded EXIF, hidden text PII, ' +
      'tracked revisions, comments, and (on request) macro/active parts through the same engine scrub ' +
      'path CLI sanitize calls. Body text is kept. Options mirror the engine scrub keys 1:1 ' +
      '(zeroRetention, purgeRSIDs, purgeAuthors, purgeExif, purgePII, acceptAllRevisions, stripComments, ' +
      'replacementAuthor, allowMacros, stripMacros); omitted options take engine defaults. ' +
      'allowMacros and stripMacros together fail with E_VALIDATION. ' +
      'Returns {ok:true,sessionId,revision,dirty,removedRSIDs,purgedAuthors,strippedExifBytes,purgedPII,bodyPII,' +
      'acceptedRevisions,strippedComments,strippedMacros,macroParts,filesTouched,droppedRogue,...} plus the ' +
      'standard redoCleared signal when a redo chain is dropped. ' +
      'Example {"sessionId":"sess_abc","stripMacros":true} -> scrubbed session at a new revision (save after). ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions.'
  ),
  normalize_document: withSharedDocs(
    'Canonical repair for an open session document: runs the same allowlist normalizer CLI normalize calls ' +
      '(verified tag renames, whitelisted default infills, schema-sequence sorts) over the live package, ' +
      'then refreshes the session view and proves the result with a read-only validation verdict. ' +
      'Output bytes match CLI normalize on the same input. When the package is still invalid afterwards, ' +
      'the session is restored untouched and the call fails with E_VALIDATION plus a fix-and-retry remedy ' +
      '(same contract as CLI normalize writing nothing on failure). ' +
      'Returns {ok:true,sessionId,revision,dirty,normalized,validation:{valid:true,...}} where normalized ' +
      'lists every rewrite once, plus the standard redoCleared signal when a redo chain is dropped. ' +
      'Example {"sessionId":"sess_abc"} -> {"ok":true,"normalized":[...]}. Save after to persist. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions.'
  ),
  import_document: withSharedDocs(
    'Ingest an on-disk Markdown file into the OPEN session (the file-ingress half compose_document lacks: ' +
      'compose takes an inline string only). Reads the file under the workspace jail, then runs the identical ' +
      'compose pipeline (strict compile gate, quotas, template-seed check, inline-span splitting), so the ' +
      'result is byte-equal to composing the same Markdown inline. sessionId is strictly required: the tool ' +
      'never auto-creates sessions (fresh templated document: new_document first, then import; existing file: ' +
      'open_document first, then import). Accepts .md/.markdown/.txt; .html/.htm refuse with E_VALIDATION ' +
      '(convert to Markdown first, same verdict as CLI import); .docx refuses with E_VALIDATION naming ' +
      'open_document (open .docx files directly). Outside-jail reads fail E_SANDBOX, missing files E_IO, ' +
      'oversize files E_LIMIT. Optional allowedDir (image refs), template (report|memo|blank seed check), ' +
      'and styleGuide (same attach semantics as compose_document). ' +
      'Optional dollarMath (default false) treats $...$ as inline math; off by default because prices and shell text can look like math. ' +
      'Example {"sessionId":"sess_abc","path":"documents/notes.md"} -> compose envelope. ' +
      'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions.'
  ),
};
