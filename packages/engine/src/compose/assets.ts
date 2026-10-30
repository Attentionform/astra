import fs from 'fs';
import path from 'path';
import { SUPPORTED_CHART_2D_AND_3D_TYPES } from '../operations/chart';
import { canonicalizeForAsset, isPathInsideDir } from '../sandbox/paths';
import { containsSvgActiveContent, containsXxeAdjacent } from '../ooxml/xxe-guard';
import {
  ALLOWED_IMAGE_EXTS,
  COMPOSE_EMU_PER_PX,
  COMPOSE_IMAGE_MAX_HEIGHT_EMU,
  COMPOSE_IMAGE_MAX_WIDTH_EMU,
  fail,
} from './constants';

// Q0a: jail primitives live in ../sandbox/paths (imported above).

export function resolveAllowedDir(allowedDir: string | undefined, line: number): string {
  if (typeof allowedDir !== 'string' || allowedDir.trim().length === 0) {
    fail(line, 'image or chart reference needs an allowed directory', 'pass allowedDir pointing at the folder that holds referenced files');
  }
  const trimmed = (allowedDir as string).trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed) || trimmed.includes('\n') || trimmed.includes('\r')) {
    fail(1, 'allowed directory path is not usable', 'pass a plain local directory path without null bytes or line breaks');
  }
  let resolved = path.resolve(trimmed);
  try {
    const stat = fs.statSync(resolved);
    if (!stat.isDirectory()) {
      fail(1, 'allowed directory is not a directory', 'pass an existing local directory that holds referenced files');
    }
  } catch {
    fail(1, 'allowed directory does not exist', 'pass an existing local directory that holds referenced files');
  }
  try {
    resolved = fs.realpathSync(resolved);
  } catch { /* keep lexical form */ }
  return resolved;
}

export function resolveAssetPath(allowedDir: string, refPath: string, line: number, what: string): string {
  if (typeof refPath !== 'string' || refPath.trim().length === 0) {
    fail(line, `${what} reference has an empty path`, 'point the reference at a file inside the allowed directory');
  }
  const trimmed = refPath.trim();
  if (trimmed.includes('\0') || /%00/i.test(trimmed) || trimmed.includes('\n') || trimmed.includes('\r')) {
    fail(line, `${what} reference path is not usable`, 'use a plain relative file path inside the allowed directory');
  }
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(trimmed)) {
    fail(line, `${what} reference points outside the allowed directory`, 'keep referenced files inside the allowed directory and use a relative path');
  }
  const candidate = path.isAbsolute(trimmed) ? path.resolve(trimmed) : path.resolve(allowedDir, trimmed);
  const canonicalAllowed = canonicalizeForAsset(allowedDir);
  const canonicalCandidate = canonicalizeForAsset(candidate);
  if (!isPathInsideDir(canonicalCandidate, canonicalAllowed) && canonicalCandidate !== canonicalAllowed) {
    fail(line, `${what} reference points outside the allowed directory`, 'keep referenced files inside the allowed directory and use a relative path');
  }
  const rel = path.relative(canonicalAllowed, canonicalCandidate);
  const segs = rel.split(path.sep).filter((s) => s && s !== '.' && s !== '..');
  for (const seg of segs) {
    if (seg.startsWith('.')) {
      fail(line, `${what} reference traverses a hidden path`, 'keep referenced files in a visible folder inside the allowed directory');
    }
  }
  const base = path.basename(canonicalCandidate);
  if (base.startsWith('.')) {
    fail(line, `${what} reference uses a hidden file name`, 'rename the file without a leading dot');
  }
  return canonicalCandidate;
}

export function detectImageKind(buf: Buffer): string | null {
  if (buf.length >= 4 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return 'png';
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'jpeg';
  if (buf.length >= 4 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x38) return 'gif';
  if (
    buf.length >= 12 &&
    buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46 &&
    buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50
  ) return 'webp';
  if (buf.length >= 2 && buf[0] === 0x42 && buf[1] === 0x4d) return 'bmp';
  if (
    buf.length >= 4 &&
    ((buf[0] === 0x49 && buf[1] === 0x49 && buf[2] === 0x2a && buf[3] === 0x00) ||
      (buf[0] === 0x4d && buf[1] === 0x4d && buf[2] === 0x00 && buf[3] === 0x2a))
  ) return 'tiff';
  const head = buf.subarray(0, 1024).toString('latin1').replace(/^\uFEFF/, '').trimStart();
  if (/^<svg[\s>/]/i.test(head)) return 'svg';
  if (/^<\?xml[\s?]/i.test(head) && /<svg[\s>/]/i.test(head.slice(0, 1024))) return 'svg';
  return null;
}

export interface ImagePixelSize {
  widthPx: number;
  heightPx: number;
}

export function validPixelSize(widthPx: number, heightPx: number): ImagePixelSize | null {
  if (!Number.isFinite(widthPx) || !Number.isFinite(heightPx)) return null;
  if (widthPx <= 0 || heightPx <= 0 || widthPx > 100000 || heightPx > 100000) return null;
  return { widthPx: Math.floor(widthPx), heightPx: Math.floor(heightPx) };
}

export function readU16(buf: Buffer, offset: number, littleEndian: boolean): number | null {
  if (offset < 0 || offset + 2 > buf.length) return null;
  return littleEndian ? buf.readUInt16LE(offset) : buf.readUInt16BE(offset);
}

export function readU32(buf: Buffer, offset: number, littleEndian: boolean): number | null {
  if (offset < 0 || offset + 4 > buf.length) return null;
  return littleEndian ? buf.readUInt32LE(offset) : buf.readUInt32BE(offset);
}

/**
 * Decodes intrinsic pixel dimensions for the image kinds the composer
 * accepts. Every branch is bounds-checked and returns null when the bytes
 * do not carry a usable size (caller then omits width/height and the engine
 * fallback applies). Never throws.
 */
export function decodeImagePixels(buf: Buffer, kind: string): ImagePixelSize | null {
  try {
    switch (kind) {
      case 'png': {
        if (buf.length < 24) return null;
        return validPixelSize(buf.readUInt32BE(16), buf.readUInt32BE(20));
      }
      case 'gif': {
        if (buf.length < 10) return null;
        return validPixelSize(buf.readUInt16LE(6), buf.readUInt16LE(8));
      }
      case 'bmp': {
        if (buf.length < 26) return null;
        const dibSize = buf.readUInt32LE(14);
        if (dibSize === 12) {
          return validPixelSize(buf.readUInt16LE(18), buf.readUInt16LE(20));
        }
        if (dibSize >= 40) {
          const w = buf.readInt32LE(18);
          const h = buf.readInt32LE(22);
          return validPixelSize(w, Math.abs(h));
        }
        return null;
      }
      case 'jpeg': {
        let offset = 2;
        while (offset + 3 < buf.length) {
          if (buf[offset] !== 0xff) {
            offset += 1;
            continue;
          }
          const marker = buf[offset + 1];
          if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
            offset += 2;
            continue;
          }
          if (marker === 0xd9 || marker === 0xda) return null;
          const startOfFrame =
            (marker >= 0xc0 && marker <= 0xc3) ||
            (marker >= 0xc5 && marker <= 0xc7) ||
            (marker >= 0xc9 && marker <= 0xcb) ||
            (marker >= 0xcd && marker <= 0xcf);
          if (startOfFrame) {
            if (offset + 9 > buf.length) return null;
            return validPixelSize(buf.readUInt16BE(offset + 7), buf.readUInt16BE(offset + 5));
          }
          if (offset + 3 > buf.length) return null;
          const length = buf.readUInt16BE(offset + 2);
          if (length < 2) return null;
          offset += 2 + length;
        }
        return null;
      }
      case 'webp': {
        if (buf.length < 16) return null;
        if (buf.subarray(0, 4).toString('ascii') !== 'RIFF') return null;
        if (buf.subarray(8, 12).toString('ascii') !== 'WEBP') return null;
        const chunk = buf.subarray(12, 16).toString('ascii');
        if (chunk === 'VP8 ') {
          if (buf.length < 30) return null;
          return validPixelSize(buf.readUInt16LE(26) & 0x3fff, buf.readUInt16LE(28) & 0x3fff);
        }
        if (chunk === 'VP8L') {
          if (buf.length < 25) return null;
          const bits = buf.readUInt32LE(21);
          return validPixelSize((bits & 0x3fff) + 1, ((bits >> 14) & 0x3fff) + 1);
        }
        if (chunk === 'VP8X') {
          if (buf.length < 30) return null;
          return validPixelSize(
            (buf[24] | (buf[25] << 8) | (buf[26] << 16)) + 1,
            (buf[27] | (buf[28] << 8) | (buf[29] << 16)) + 1
          );
        }
        return null;
      }
      case 'tiff': {
        if (buf.length < 8) return null;
        const littleEndian = buf[0] === 0x49 && buf[1] === 0x49;
        const bigEndian = buf[0] === 0x4d && buf[1] === 0x4d;
        if (!littleEndian && !bigEndian) return null;
        const magic = readU16(buf, 2, littleEndian);
        if (magic !== 42) return null;
        const ifdOffset = readU32(buf, 4, littleEndian);
        if (ifdOffset === null || ifdOffset < 8 || ifdOffset + 2 > buf.length) return null;
        const count = readU16(buf, ifdOffset, littleEndian);
        if (count === null || count <= 0 || count > 200) return null;
        let width: number | null = null;
        let height: number | null = null;
        for (let i = 0; i < count; i++) {
          const entry = ifdOffset + 2 + i * 12;
          if (entry + 12 > buf.length) break;
          const tag = readU16(buf, entry, littleEndian);
          const fieldType = readU16(buf, entry + 2, littleEndian);
          const fieldCount = readU32(buf, entry + 4, littleEndian);
          if (tag === null || fieldType === null || fieldCount === null) continue;
          if ((tag !== 256 && tag !== 257) || fieldCount !== 1) continue;
          let value: number | null = null;
          if (fieldType === 3) {
            value = readU16(buf, entry + 8, littleEndian);
          } else if (fieldType === 4) {
            value = readU32(buf, entry + 8, littleEndian);
          } else {
            continue;
          }
          if (value === null) continue;
          if (tag === 256) width = value;
          else height = value;
        }
        if (width === null || height === null) return null;
        return validPixelSize(width, height);
      }
      case 'svg': {
        return decodeSvgPixels(buf);
      }
      default:
        return null;
    }
  } catch {
    return null;
  }
}

export function svgLengthToPx(raw: string): number | null {
  const match = /^\s*([0-9]*\.?[0-9]+(?:[eE][+-]?[0-9]+)?)\s*(px|pt|pc|mm|cm|q|in)?\s*$/i.exec(raw);
  if (!match) return null;
  const value = Number(match[1]);
  if (!Number.isFinite(value) || value <= 0) return null;
  const unit = (match[2] ?? 'px').toLowerCase();
  const perUnit: Record<string, number> = {
    px: 1,
    in: 96,
    cm: 96 / 2.54,
    mm: 96 / 25.4,
    q: 96 / 254,
    pt: 96 / 72,
    pc: 16,
  };
  const factor = perUnit[unit];
  if (factor === undefined) return null;
  return value * factor;
}

/** Reads SVG width/height attributes, falling back to the viewBox size. */
export function decodeSvgPixels(buf: Buffer): ImagePixelSize | null {
  try {
    const text = buf.subarray(0, 65536).toString('utf8');
    const tagMatch = /<svg[\s>][^>]*>?/i.exec(text);
    if (!tagMatch) return null;
    const tag = tagMatch[0];
    const attr = (name: string): string | null => {
      const found = new RegExp(`${name}\\s*=\\s*["']([^"']+)["']`, 'i').exec(tag);
      return found ? found[1] : null;
    };
    const rawWidth = attr('width');
    const rawHeight = attr('height');
    if (rawWidth && rawHeight) {
      const w = svgLengthToPx(rawWidth);
      const h = svgLengthToPx(rawHeight);
      if (w !== null && h !== null) return validPixelSize(w, h);
    }
    const rawBox = attr('viewBox');
    if (rawBox) {
      const parts = rawBox.trim().split(/[\s,]+/).map(Number);
      if (parts.length === 4 && parts.every((n) => Number.isFinite(n))) {
        return validPixelSize(parts[2], parts[3]);
      }
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Maps intrinsic pixels to EMU drawing sizes that keep the aspect ratio and
 * fit the page content box. Images smaller than the box keep their intrinsic
 * size (scale never exceeds 1); larger ones shrink until both caps hold.
 * Returns null when no usable size was decoded (engine fallback applies).
 */
export function fitImageEmus(px: ImagePixelSize | null): { width: number; height: number } | null {
  if (!px) return null;
  const rawW = px.widthPx * COMPOSE_EMU_PER_PX;
  const rawH = px.heightPx * COMPOSE_EMU_PER_PX;
  if (!Number.isFinite(rawW) || !Number.isFinite(rawH) || rawW <= 0 || rawH <= 0) return null;
  const scale = Math.min(1, COMPOSE_IMAGE_MAX_WIDTH_EMU / rawW, COMPOSE_IMAGE_MAX_HEIGHT_EMU / rawH);
  return {
    width: Math.max(1, Math.round(rawW * scale)),
    height: Math.max(1, Math.round(rawH * scale)),
  };
}

/**
 * H2/D4 fd-hardened asset read (best-effort + documented residual).
 *
 * Closes the stat→read TOCTOU that a plain `statSync`+`readFileSync` pair
 * leaves open: the path is re-canonicalized pre-open, opened with
 * `O_NOFOLLOW` where the platform offers it (symlink opens fail with ELOOP
 * instead of following), sized via `fstat` on the OPEN fd (never the path),
 * read through the fd, then re-checked post-read (a swap between open and
 * read still fails closed here). Failure details/remedies are caller-
 * supplied so image vs chart error texts stay byte-identical to before.
 *
 * Documented residual (D4, best-effort accepted): on platforms without
 * `O_NOFOLLOW`, and against a privileged local attacker racing the open
 * itself, the window narrows to the single open syscall but does not close
 * — that attacker already owns the box in every way that matters (it can
 * read the files directly), so refusing the operation there would trade
 * real availability for no meaningful protection.
 */
export interface AssetReadTexts {
  /** Noun for messages (`image` / `chart`). */
  what: string;
  missingRemedy: string;
  notFileDetail: string;
  notFileRemedy: string;
  emptyDetail: string;
  emptyRemedy: string;
  tooLargeDetail: string;
  tooLargeRemedy: string;
  unreadableDetail: string;
  unreadableRemedy: string;
}

export function readAssetBytesHardened(
  absPath: string,
  allowedDir: string,
  line: number,
  texts: AssetReadTexts,
  maxBytes: number
): Buffer {
  const base = path.basename(absPath);
  const canonicalAllowed = canonicalizeForAsset(allowedDir);
  const pre = canonicalizeForAsset(absPath);
  if (pre !== canonicalAllowed && !isPathInsideDir(pre, canonicalAllowed)) {
    fail(line, `${texts.what} reference points outside the allowed directory`, 'keep referenced files inside the allowed directory and use a relative path');
  }
  const nofollow = fs.constants.O_NOFOLLOW as number | undefined;
  const flags =
    typeof nofollow === 'number' && nofollow !== 0
      ? (fs.constants.O_RDONLY as number) | nofollow
      : 'r';
  let fd: number | undefined;
  try {
    fd = fs.openSync(pre, flags);
  } catch (err: unknown) {
    if ((err as { code?: string } | undefined)?.code === 'ENOENT') {
      fail(line, `${texts.what} file is missing (${base})`, texts.missingRemedy);
    }
    fail(line, texts.unreadableDetail, texts.unreadableRemedy);
  }
  try {
    const st = fs.fstatSync(fd as number);
    if (!st.isFile()) {
      fail(line, texts.notFileDetail, texts.notFileRemedy);
    }
    if (st.size <= 0) {
      fail(line, texts.emptyDetail, texts.emptyRemedy);
    }
    if (st.size > maxBytes) {
      fail(line, texts.tooLargeDetail, texts.tooLargeRemedy);
    }
    const buf = fs.readFileSync(fd as number);
    if (buf.length > maxBytes) {
      fail(line, texts.tooLargeDetail, texts.tooLargeRemedy);
    }
    let post = pre;
    try {
      post = fs.realpathSync(absPath);
    } catch {
      /* target vanished mid-read — pre-open verdict stands */
    }
    if (post !== canonicalAllowed && !isPathInsideDir(post, canonicalAllowed)) {
      fail(line, `${texts.what} reference points outside the allowed directory`, 'keep referenced files inside the allowed directory and use a relative path');
    }
    return buf;
  } finally {
    try {
      if (fd !== undefined) fs.closeSync(fd);
    } catch {
      /* close must never mask the read verdict */
    }
  }
}

/**
 * H3/D5 compose-ingest SVG gate (refuse path): hostile SVG constructs fail
 * the whole compose with E_VALIDATION + line number. Covers the legacy set
 * (doctype/entity/script) plus the adjacent vectors (xi:include, xsl:,
 * xml-stylesheet PI) and SVG script-capable shapes (event-handler
 * attributes, foreignObject). Legit SVGs (plain shapes, width/height or
 * viewBox for sizing) pass untouched.
 */
export function assertSvgAssetClean(text: string, line: number, base: string): void {
  if (/<!DOCTYPE|<!ENTITY/i.test(text) || containsXxeAdjacent(text) || containsSvgActiveContent(text)) {
    fail(
      line,
      `image SVG carries active or entity markup (${base})`,
      'use a plain SVG without doctype, entities, script, event attributes, or javascript: hrefs (plain <style> blocks pass ingest and are stripped with report on sanitize)'
    );
  }
  if (/<script/i.test(text)) {
    fail(
      line,
      `image SVG carries active or entity markup (${base})`,
      'use a plain SVG without doctype, entities, script, event attributes, or javascript: hrefs (plain <style> blocks pass ingest and are stripped with report on sanitize)'
    );
  }
}

export function readImageAsset(absPath: string, line: number, allowedDir: string): { dataUrl: string; width?: number; height?: number } {
  const ext = path.extname(absPath).toLowerCase().replace(/^\./, '');
  if (!ALLOWED_IMAGE_EXTS.has(ext)) {
    fail(line, `image file extension .${ext || '(none)'} is not supported`, 'use png, jpg, gif, webp, bmp, tiff, or svg');
  }
  const buf = readAssetBytesHardened(absPath, allowedDir, line,
    {
      what: 'image',
      missingRemedy: 'add the file inside the allowed directory or fix the reference path',
      notFileDetail: `image reference is not a file (${path.basename(absPath)})`,
      notFileRemedy: 'point the reference at a regular image file',
      emptyDetail: `image file is empty (${path.basename(absPath)})`,
      emptyRemedy: 'use a non-empty image file',
      tooLargeDetail: `image file is too large (${path.basename(absPath)})`,
      tooLargeRemedy: 'use a smaller image file',
      unreadableDetail: `image file cannot be read (${path.basename(absPath)})`,
      unreadableRemedy: 'check the file exists and is readable, then retry',
    },
    15_000_000);
  const kind = detectImageKind(buf);
  if (!kind) {
    fail(line, `image file bytes are not a known image (${path.basename(absPath)})`, 'use a real png, jpg, gif, webp, bmp, tiff, or svg file');
  }
  if (kind === 'svg') {
    assertSvgAssetClean(buf.toString('utf8'), line, path.basename(absPath));
  }
  const mime = kind === 'jpeg' ? 'image/jpeg' : kind === 'svg' ? 'image/svg+xml' : kind === 'tiff' ? 'image/tiff' : `image/${kind}`;
  const fitted = fitImageEmus(decodeImagePixels(buf, kind as string));
  if (!fitted) {
    return { dataUrl: `data:${mime};base64,${buf.toString('base64')}` };
  }
  return { dataUrl: `data:${mime};base64,${buf.toString('base64')}`, width: fitted.width, height: fitted.height };
}

export interface ChartAsset {
  chartType: string;
  title?: string;
  categories?: string[];
  series?: Array<{ name: string; values: number[]; color?: string }>;
}

export function readChartAsset(absPath: string, line: number, allowedDir: string): ChartAsset {
  const raw = readAssetBytesHardened(absPath, allowedDir, line,
    {
      what: 'chart',
      missingRemedy: 'add the JSON file inside the allowed directory or fix the reference path',
      notFileDetail: 'chart reference is not a file',
      notFileRemedy: 'point the reference at a regular JSON file',
      emptyDetail: 'chart file is empty',
      emptyRemedy: 'write chart JSON with title, categories, and series',
      tooLargeDetail: 'chart file is too large',
      tooLargeRemedy: 'keep chart JSON under one megabyte',
      unreadableDetail: 'chart file cannot be read',
      unreadableRemedy: 'check the file exists and is readable, then retry',
    },
    1_000_000).toString('utf8');
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    fail(line, 'chart file is not valid JSON', 'write an object with title, categories, and series');
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    fail(line, 'chart file must hold a JSON object', 'write an object with title, categories, and series');
  }
  const rec = parsed as Record<string, unknown>;
  const chartType = typeof rec['chartType'] === 'string' && rec['chartType'].trim().length > 0 ? rec['chartType'].trim() : 'column';
  if (!(SUPPORTED_CHART_2D_AND_3D_TYPES as readonly string[]).includes(chartType)) {
    fail(line, `chart type "${chartType}" is not supported`, `use one of ${(SUPPORTED_CHART_2D_AND_3D_TYPES as readonly string[]).join(', ')}`);
  }
  let title: string | undefined;
  if (rec['title'] !== undefined) {
    if (typeof rec['title'] !== 'string' || rec['title'].trim().length === 0) {
      fail(line, 'chart title must be a non-empty string', 'pass a title or omit it');
    }
    title = (rec['title'] as string).trim();
    if (title.length > 1000) {
      fail(line, 'chart title is too long', 'keep the title under 1000 characters');
    }
  }
  let categories: string[] | undefined;
  if (rec['categories'] !== undefined) {
    if (!Array.isArray(rec['categories']) || rec['categories'].length === 0) {
      fail(line, 'chart categories must be a non-empty list', 'pass categories as a list of labels');
    }
    categories = (rec['categories'] as unknown[]).map((entry, idx) => {
      if (typeof entry !== 'string' || entry.trim().length === 0) {
        fail(line, `chart category ${idx + 1} must be a non-empty string`, 'use plain labels for categories');
      }
      return (entry as string).trim();
    }) as string[];
    if ((categories as string[]).length > 1000) {
      fail(line, 'chart has too many categories', 'keep categories at 1000 or fewer');
    }
  }
  let series: ChartAsset['series'];
  if (rec['series'] !== undefined) {
    if (!Array.isArray(rec['series']) || rec['series'].length === 0) {
      fail(line, 'chart series must be a non-empty list', 'pass series as a list of {name, values}');
    }
    series = (rec['series'] as unknown[]).map((entry, idx) => {
      if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
        fail(line, `chart series ${idx + 1} must be an object`, 'use {name, values} per series');
      }
      const srec = entry as Record<string, unknown>;
      if (typeof srec['name'] !== 'string' || (srec['name'] as string).trim().length === 0) {
        fail(line, `chart series ${idx + 1} needs a non-empty name`, 'give each series a name');
      }
      if (!Array.isArray(srec['values']) || (srec['values'] as unknown[]).length === 0) {
        fail(line, `chart series ${idx + 1} needs non-empty values`, 'give each series a list of numbers');
      }
      const values = (srec['values'] as unknown[]).map((v, vi) => {
        if (typeof v !== 'number' || !Number.isFinite(v)) {
          fail(line, `chart series ${idx + 1} value ${vi + 1} must be a number`, 'use finite numbers for series values');
        }
        return v as number;
      });
      const out: { name: string; values: number[]; color?: string } = {
        name: (srec['name'] as string).trim(),
        values,
      };
      if (srec['color'] !== undefined) {
        if (typeof srec['color'] !== 'string' || (srec['color'] as string).trim().length === 0) {
          fail(line, `chart series ${idx + 1} color must be a non-empty string`, 'omit color or pass a hex string');
        }
        out.color = (srec['color'] as string).trim();
      }
      return out;
    });
    if ((series as Array<unknown>).length > 100) {
      fail(line, 'chart has too many series', 'keep series at 100 or fewer');
    }
  }
  return { chartType, title, categories, series };
}
