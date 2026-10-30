// ── S1b / EX-05: IPC schema + size caps ───────────────────────────────────────
// Renderer-controlled IPC payloads (send-chat / execute-operations /
// update-node) previously had zero schema/size validation: 50MB messages →
// OOM, 10k-op batches → event-loop stall, spoofable modelId.
//
// DECISION (plan §5.2 — "no 50MB hard cap"): caps are sized to reject
// OOM/stall-scale abuse while letting legitimate media documents through.
// Text and media are budgeted SEPARATELY (media exemption): vision
// screenshots (~70k tokens, multi-MB base64 data URLs) must PASS even though
// the equivalent bytes of plain text would be rejected. Rationale: base64
// media is opaque to the event loop (never parsed as operations/XML) and is
// already bounded by its own media cap; text IS parsed (JSON/XML/engine) and
// gets the tight cap.
//
// Caps (documented, tunable):
//   text  ≤ 1MB   (JSON text with data-URLs stripped)
//   media ≤ 15MB  (sum of data:...;base64, segments)
//   messages ≤ 200, operations ≤ 500, single op ≤ 200KB text, node XML ≤ 500KB
// No 50MB cap anywhere by design.
//
// Schema allowlisting is EXTENSIBLE: SEND_CHAT_ALLOWED_KEYS includes the U1
// future `temperature` field now, so U1 plumbing does not trip this gate.
// Unknown top-level keys are rejected (fail-closed) — add future fields here.
//
// Lives here (not electron/main.ts) so vitest can import it without loading
// Electron. electron/main.ts handlers MUST call these validators first.
//
// ── Track-B: strict operation schema defense ────────────────────────────────
// `validateExecuteOperationsPayload` below enforces the per-action
// `z.discriminatedUnion` (@attentionform/astra-engine operation-schema) AFTER the cheap
// structural/size gates. Unknown actions, missing required fields, and
// unknown flat keys throw E_VALIDATION with a remedy — model garbage is
// rejected at the IPC gate, not after wasted engine steps. Returns ok:true
// via assertion (throws otherwise); never a silent fallback, never null.
import { validateOperationsStrict, type Operation } from './operation-schema';

export const S1B_MAX_CHAT_TEXT_BYTES = 1_000_000;
export const S1B_MAX_CHAT_MEDIA_BYTES = 15_000_000;
export const S1B_MAX_CHAT_MESSAGES = 200;
export const S1B_MAX_MESSAGE_TEXT_BYTES = 500_000;
export const S1B_MAX_OPERATIONS = 500;
export const S1B_MAX_OPERATION_TEXT_BYTES = 200_000;
export const S1B_MAX_NODE_XML_BYTES = 500_000;
export const S1B_MAX_MODEL_ID_LENGTH = 128;
export const S1B_MAX_NODE_ID_LENGTH = 256;
export const S1B_MAX_REQUEST_ID_LENGTH = 128;

// U1 EXTENSION POINT: add future send-chat fields here (NOT inline in main.ts).
export const S1B_SEND_CHAT_ALLOWED_KEYS = new Set([
  'messages',
  'modelId',
  'requestId',
  'thinkingLevel',
  // U1 future field — accepted now so temperature plumbing never trips S1b.
  'temperature',
]);

export const S1B_THINKING_LEVELS = new Set(['off', 'low', 'medium', 'high']);

const MODEL_ID_PATTERN = /^[A-Za-z0-9._:/@-]+$/;
// Matches data URLs carrying base64 payloads (screenshots, images, …).
// NOTE: the media EXEMPTION is narrower than this match — see
// S1B_ALLOWED_MEDIA_TOP_LEVELS / isAllowedMediaMime below. A data URL whose
// MIME is not allowlisted (notably text/*) is NOT media: it stays in the text
// budget (S1b follow-up item 3 — data:text/plain smuggling closure).
const DATA_URL_RE = /data:([a-zA-Z0-9.+-]+)\/([a-zA-Z0-9.+-]+)((?:;[a-zA-Z0-9.+-=]+)*);base64,([A-Za-z0-9+/=]+)/g;

// ── S1b follow-up item 3: media-exemption MIME allowlist ─────────────────────
// What legitimately flows through these validators as data URLs (verified by
// code inspection, Sept 2026):
//   visual_review  → image/jpeg (visual-review.ts captureDocumentPages) and
//                     image/png (NativeCanvasViewer screenshot path)
//   read_local_file → image/png, image/jpeg, image/webp (local-file-reader.ts;
//                     .gif/.bmp inputs are also images)
//   fetch_web_image → image/* (agent.ts enforces content-type image/*)
//   image ops       → data:image/* only (operations/image.ts DATA_URL_PATTERN)
//   parser media    → image/* (parser.ts getMimeType)
// Chat attachments send only "[Attached Document: path]" TEXT (ChatPanel.tsx)
// — never a data URL — so no text/* or application/* sender exists in-tree.
// Allowlist: image/*, video/*, audio/*, application/pdf (common legit media;
// video/audio/pdf are permitted for forward-compat, still bounded by the 15MB
// media cap). Everything else — ESPECIALLY text/* (counts as TEXT) — stays in
// the text budget, so multi-MB data:text/plain can no longer ride the media
// exemption past the 1MB text cap.
const S1B_ALLOWED_MEDIA_TOP_LEVELS = new Set(['image', 'video', 'audio']);
const S1B_ALLOWED_MEDIA_EXACT = new Set(['application/pdf']);

/** True when a data-URL MIME qualifies for the media exemption (not text). */
export function isAllowedMediaMime(topLevel: string, subType: string): boolean {
  const top = (topLevel || '').toLowerCase();
  const sub = (subType || '').toLowerCase();
  if (S1B_ALLOWED_MEDIA_TOP_LEVELS.has(top)) return true;
  if (S1B_ALLOWED_MEDIA_EXACT.has(`${top}/${sub}`)) return true;
  return false;
}

// ── S1b R3 item 1: MIME label-spoof closure — magic-byte sniff ───────────────
// A data:image/png LABEL alone is not proof of image bytes: plain-text base64
// relabeled as image/png rode the 15MB media exemption past the 1MB text cap
// (verified by execution probe, Sept 2026: 4MB 'A's labeled image/png →
// mediaBytes 4M, textBytes ~7, PASS). Fix: sniff the first decoded bytes
// (O(few bytes) — only the first ~64 base64 chars are decoded, never the full
// multi-MB body) and require the magic to match the claimed MIME.
// Mismatch (or allowlisted-but-unverifiable subtype with no known signature)
// → counts as TEXT (fail-closed to the text budget, NOT a reject — small legit
// edge types like a tiny BMP still pass via text; large spoofs hit the text
// cap). Exemption is kept for sniff-VERIFIED PNG/JPEG/GIF/WebP/PDF only.
// Perf: slice+decode ≤64 chars ≈ microseconds; measured <2ms for a 4MB
// payload on this box (see S1b test "media sniff is O(few bytes)").
// Residual: video/*, audio/* and exotic image/* (bmp/tiff/avif/…) have no
// signature checked here and therefore count as TEXT by design (fail-closed).
// No in-tree sender emits them as data URLs (see allowlist comment above);
// small ones still pass via the text budget, large ones correctly trip it.
const S1B_SNIFF_HEAD_B64_CHARS = 64;
const S1B_JPEG_SUBS = new Set(['jpeg', 'jpg', 'jpe', 'jfif', 'pjpeg', 'pjp']);

function decodeSniffHead(b64: string): Buffer | null {
  try {
    if (!b64) return null;
    const head = b64.slice(0, S1B_SNIFF_HEAD_B64_CHARS);
    if (!head) return null;
    const buf = Buffer.from(head, 'base64');
    return buf && buf.length > 0 ? buf : null;
  } catch {
    return null;
  }
}

/**
 * True only when the claimed media MIME is allowlisted AND the payload's
 * magic bytes match (sniff-verified). Used for the media exemption decision.
 * Fail-closed: unknown/unverifiable subtypes and decode failures → false
 * (counts as TEXT).
 */
export function isSniffVerifiedMedia(topLevel: string, subType: string, b64Payload: string): boolean {
  if (!isAllowedMediaMime(topLevel, subType)) return false;
  const top = (topLevel || '').toLowerCase();
  const sub = (subType || '').toLowerCase();
  const head = decodeSniffHead(b64Payload || '');
  if (!head || head.length < 3) return false;
  // PNG: 89 50 4E 47 0D 0A 1A 0A
  if (top === 'image' && (sub === 'png' || sub === 'x-png')) {
    return head.length >= 8 &&
      head[0] === 0x89 && head[1] === 0x50 && head[2] === 0x4e && head[3] === 0x47 &&
      head[4] === 0x0d && head[5] === 0x0a && head[6] === 0x1a && head[7] === 0x0a;
  }
  // JPEG: FF D8 FF
  if (top === 'image' && S1B_JPEG_SUBS.has(sub)) {
    return head[0] === 0xff && head[1] === 0xd8 && head[2] === 0xff;
  }
  // GIF: "GIF87a" | "GIF89a" (47 49 46 38 37/39 61)
  if (top === 'image' && sub === 'gif') {
    return head.length >= 6 &&
      head[0] === 0x47 && head[1] === 0x49 && head[2] === 0x46 && head[3] === 0x38 &&
      (head[4] === 0x37 || head[4] === 0x39) && head[5] === 0x61;
  }
  // WebP: "RIFF" xxxx "WEBP" (52 49 46 46 … 57 45 42 50 at offset 8)
  if (top === 'image' && sub === 'webp') {
    return head.length >= 12 &&
      head[0] === 0x52 && head[1] === 0x49 && head[2] === 0x46 && head[3] === 0x46 &&
      head[8] === 0x57 && head[9] === 0x45 && head[10] === 0x42 && head[11] === 0x50;
  }
  // PDF: "%PDF" (25 50 44 46), 5th byte typically 2D ("-")
  if (top === 'application' && sub === 'pdf') {
    return head.length >= 4 &&
      head[0] === 0x25 && head[1] === 0x50 && head[2] === 0x44 && head[3] === 0x46;
  }
  // Allowlisted but no known signature (video/*, audio/*, image/bmp, …) →
  // fail-closed to TEXT (not exempt). See residual note above.
  return false;
}

export function stripDataUrls(s: string): string {
  if (typeof s !== 'string' || !s) return typeof s === 'string' ? s : '';
  // Strip ONLY sniff-verified media URLs; label-spoofed (mismatched magic) and
  // non-media data URLs (text/*, …) stay as text toward the text cap (R3 item 1).
  return s.replace(DATA_URL_RE, (m, top, sub, _params, b64) =>
    isSniffVerifiedMedia(top, sub, b64) ? '' : m
  );
}

export function countMediaBytes(s: string): number {
  if (typeof s !== 'string' || !s) return 0;
  let total = 0;
  DATA_URL_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = DATA_URL_RE.exec(s)) !== null) {
    if (isSniffVerifiedMedia(m[1], m[2], m[4])) total += Buffer.byteLength(m[0], 'utf8');
  }
  DATA_URL_RE.lastIndex = 0;
  return total;
}

/** Byte sizes of a JSON-serializable value, split into text vs media. */
export function measureTextAndMedia(value: unknown): { textBytes: number; mediaBytes: number } {
  let serialized: string;
  try {
    serialized = JSON.stringify(value) ?? '';
  } catch {
    throw new Error('IPC payload is not JSON-serializable');
  }
  const mediaBytes = countMediaBytes(serialized);
  const textBytes = Buffer.byteLength(stripDataUrls(serialized), 'utf8');
  return { textBytes, mediaBytes };
}

export function validateModelId(modelId: unknown): asserts modelId is string {
  if (typeof modelId !== 'string' || !modelId.trim()) {
    throw new Error('IPC validation: "modelId" must be a non-empty string');
  }
  const id = modelId.trim();
  if (id.length > S1B_MAX_MODEL_ID_LENGTH) {
    throw new Error(`IPC validation: "modelId" exceeds ${S1B_MAX_MODEL_ID_LENGTH} chars`);
  }
  // Path/URL smuggling has no place in a model id (legit ids are
  // provider/model tokens like openrouter/meta-llama-x, ollama/y, gpt-4o-mini).
  if (id.includes('..') || id.includes('\0') || /\s/.test(id)) {
    throw new Error('IPC validation: "modelId" contains forbidden characters');
  }
  if (!MODEL_ID_PATTERN.test(id)) {
    throw new Error('IPC validation: "modelId" contains forbidden characters');
  }
}

function validateThinkingLevel(value: unknown): void {
  if (value === undefined) return;
  if (typeof value !== 'string' || !S1B_THINKING_LEVELS.has(value)) {
    throw new Error('IPC validation: "thinkingLevel" must be one of off|low|medium|high');
  }
}

function validateTemperature(value: unknown): void {
  if (value === undefined) return;
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 2) {
    throw new Error('IPC validation: "temperature" must be a number in [0, 2]');
  }
}

export interface SendChatPayload {
  messages: unknown;
  modelId: unknown;
  requestId?: unknown;
  thinkingLevel?: unknown;
  temperature?: unknown;
  [k: string]: unknown;
}

/**
 * Validates a send-chat IPC payload. Throws on violation.
 * Media exemption: data-URL segments count toward the media budget, not text.
 */
export function validateSendChatPayload(payload: SendChatPayload): void {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('IPC validation: send-chat payload must be an object');
  }
  for (const key of Object.keys(payload)) {
    if (!S1B_SEND_CHAT_ALLOWED_KEYS.has(key)) {
      throw new Error(`IPC validation: send-chat unknown field "${key}"`);
    }
  }
  const { messages, modelId, requestId, thinkingLevel, temperature } = payload;
  validateModelId(modelId);
  validateThinkingLevel(thinkingLevel);
  validateTemperature(temperature);
  if (requestId !== undefined && (typeof requestId !== 'string' || requestId.length > S1B_MAX_REQUEST_ID_LENGTH)) {
    throw new Error('IPC validation: "requestId" must be a string ≤ 128 chars');
  }
  if (!Array.isArray(messages) || messages.length === 0) {
    throw new Error('IPC validation: "messages" must be a non-empty array');
  }
  if (messages.length > S1B_MAX_CHAT_MESSAGES) {
    throw new Error(`IPC validation: "messages" exceeds ${S1B_MAX_CHAT_MESSAGES} entries`);
  }
  for (let i = 0; i < messages.length; i++) {
    const msg: unknown = messages[i];
    if (!msg || typeof msg !== 'object') {
      throw new Error(`IPC validation: messages[${i}] must be an object`);
    }
    const role: unknown = (msg as Record<string, unknown>)['role'];
    if (role !== undefined && typeof role !== 'string') {
      throw new Error(`IPC validation: messages[${i}].role must be a string`);
    }
    const { textBytes, mediaBytes } = measureTextAndMedia(msg);
    if (textBytes > S1B_MAX_MESSAGE_TEXT_BYTES) {
      throw new Error(`IPC validation: messages[${i}] text exceeds ${S1B_MAX_MESSAGE_TEXT_BYTES} bytes`);
    }
    if (mediaBytes > S1B_MAX_CHAT_MEDIA_BYTES) {
      throw new Error(`IPC validation: messages[${i}] media exceeds ${S1B_MAX_CHAT_MEDIA_BYTES} bytes`);
    }
  }
  const total = measureTextAndMedia(messages);
  if (total.textBytes > S1B_MAX_CHAT_TEXT_BYTES) {
    throw new Error(`IPC validation: chat text exceeds ${S1B_MAX_CHAT_TEXT_BYTES} bytes`);
  }
  if (total.mediaBytes > S1B_MAX_CHAT_MEDIA_BYTES) {
    throw new Error(`IPC validation: chat media exceeds ${S1B_MAX_CHAT_MEDIA_BYTES} bytes`);
  }
}

/**
 * Validates an execute-operations IPC payload. Throws on violation.
 * Image/chart ops legitimately embed base64 fileData — those segments count
 * toward the media budget (exemption), not the per-op text cap.
 * Strict defense: after the size gates, every op must pass the per-action
 * discriminatedUnion (E_VALIDATION with remedy otherwise). Asserts ok:true.
 */
export function validateExecuteOperationsPayload(operations: unknown): asserts operations is Operation[] {
  if (!Array.isArray(operations) || operations.length === 0) {
    throw new Error('IPC validation: "operations" must be a non-empty array');
  }
  if (operations.length > S1B_MAX_OPERATIONS) {
    throw new Error(`IPC validation: "operations" exceeds ${S1B_MAX_OPERATIONS} entries (stall guard)`);
  }
  for (let i = 0; i < operations.length; i++) {
    const op: unknown = operations[i];
    if (!op || typeof op !== 'object' || Array.isArray(op)) {
      throw new Error(`IPC validation: operations[${i}] must be an object`);
    }
    const record = op as Record<string, unknown>;
    const action: unknown = record['action'];
    if (typeof action !== 'string' || !action.trim()) {
      throw new Error(`IPC validation: operations[${i}].action must be a non-empty string`);
    }
    if (action.length > 64) {
      throw new Error(`IPC validation: operations[${i}].action exceeds 64 chars`);
    }
    const { textBytes, mediaBytes } = measureTextAndMedia(op);
    if (textBytes > S1B_MAX_OPERATION_TEXT_BYTES) {
      throw new Error(`IPC validation: operations[${i}] text exceeds ${S1B_MAX_OPERATION_TEXT_BYTES} bytes`);
    }
    if (mediaBytes > S1B_MAX_CHAT_MEDIA_BYTES) {
      throw new Error(`IPC validation: operations[${i}] media exceeds ${S1B_MAX_CHAT_MEDIA_BYTES} bytes`);
    }
  }
  // Track-B strict schema gate: per-action discriminatedUnion (throws
  // E_VALIDATION with remedy on unknown actions / missing fields / bad keys).
  validateOperationsStrict(operations);
  const total = measureTextAndMedia(operations);
  if (total.textBytes > S1B_MAX_CHAT_TEXT_BYTES) {
    throw new Error(`IPC validation: operations text exceeds ${S1B_MAX_CHAT_TEXT_BYTES} bytes`);
  }
  if (total.mediaBytes > S1B_MAX_CHAT_MEDIA_BYTES) {
    throw new Error(`IPC validation: operations media exceeds ${S1B_MAX_CHAT_MEDIA_BYTES} bytes`);
  }
}

/** Validates an update-node IPC payload. Throws on violation. */
export function validateUpdateNodePayload(nodeId: unknown, newXml: unknown): void {
  if (typeof nodeId !== 'string' || !nodeId.trim()) {
    throw new Error('IPC validation: "nodeId" must be a non-empty string');
  }
  if (nodeId.length > S1B_MAX_NODE_ID_LENGTH) {
    throw new Error(`IPC validation: "nodeId" exceeds ${S1B_MAX_NODE_ID_LENGTH} chars`);
  }
  if (typeof newXml !== 'string' || !newXml) {
    throw new Error('IPC validation: "newXml" must be a non-empty string');
  }
  const { textBytes } = measureTextAndMedia(newXml);
  if (textBytes > S1B_MAX_NODE_XML_BYTES) {
    throw new Error(`IPC validation: "newXml" exceeds ${S1B_MAX_NODE_XML_BYTES} bytes`);
  }
}
