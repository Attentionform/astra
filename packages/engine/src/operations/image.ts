import { DocumentAST, Node } from '../ast';
import { BaseOperation, validateVersion } from './base';
import { RelationshipsManipulator } from '../manipulators/relationships';
import { DocxParser } from '../parser';
import { insertIntoSection, lastSection, locateNode } from '../parser/sections';
import {
  createWordElement,
  createWordprocessingDrawingElement,
  createRelationshipsElement,
  createContentTypesElement,
  allocateNextDrawingDocPrId,
} from '../ooxml/ns';
import * as crypto from 'crypto';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { resolveReproducibleMode } from '../zip-hardening';
import { toError } from '../errors';
import { readImageFileBytes } from './image-file';
import { decodeImagePixels, fitImageEmus } from '../compose/assets';

export type ImageWrapMode = 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behindText' | 'inFrontOfText';

export interface ImageOperation extends BaseOperation {
  action: 'image';
  operation: 'insert' | 'resize' | 'reposition' | 'setWrap' | 'setCaption' | 'delete' | 'setAltText';
  anchor?: string; // Required for insert
  target?: string; // Required for resize, reposition, delete (targets the drawing ID)
  fileData?: string; // base64 or data URL (for insert)
  filePath?: string; // file inside allowedDir, resolved against the session jail (for insert; exclusive with fileData)
  width?: number; // In EMUs (914400 = 1 inch)
  height?: number; // In EMUs
  wrapMode?: ImageWrapMode;
  captionText?: string;
  altText?: string;
}

// Strips the data-URL prefix, e.g. "data:image/png;base64," (case-insensitive scheme).
const DATA_URL_PATTERN = /^data:image\/([\w.+-]+);base64,/i;

// Module-level counters: opt-out wall-clock collision guard only.
// Deterministic default path never consults these (content-seeded,
// AST-scoped collision loop instead) so same input twice stays identical.
let fileNameSeq = 0;
let nodeIdSeq = 0;

/** Deterministic by default; wall-clock only when opted out. */
function isImageReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

/** Content-seeded 8-hex from image bytes + dims (stable across runs). */
function imageHash8(buf: Buffer, ext: string, w: number, h: number): string {
  try {
    return crypto.createHash('sha256').update(buf).update(`:${ext}:${w}x${h}`).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

/** Collect live node ids for AST-scoped deterministic collision loops. */
function collectLiveIds(ast: DocumentAST): Set<string> {
  const live = new Set<string>();
  try {
    for (const sec of ast.sections ?? []) {
      for (const n of sec.nodes ?? []) {
        if (typeof n.id === 'string' && n.id) live.add(n.id);
      }
    }
    for (const k of Object.keys(ast.mediaFiles ?? {})) live.add(k);
  } catch { /* best-effort */ }
  return live;
}

function deriveImageFormat(mimeSubtype: string): { ext: string; contentType: string } {
  switch ((mimeSubtype || '').toLowerCase().replace(/^\./, '')) {
    case 'png':
      return { ext: 'png', contentType: 'image/png' };
    case 'jpeg':
    case 'jpg':
      return { ext: 'jpeg', contentType: 'image/jpeg' };
    case 'gif':
      return { ext: 'gif', contentType: 'image/gif' };
    case 'webp':
      return { ext: 'webp', contentType: 'image/webp' };
    case 'bmp':
      return { ext: 'bmp', contentType: 'image/bmp' };
    case 'tif':
    case 'tiff':
      return { ext: 'tiff', contentType: 'image/tiff' };
    case 'svg':
    case 'svg+xml':
      return { ext: 'svg', contentType: 'image/svg+xml' };
    default:
      // OP-08: unknown/missing format defaults to PNG (Word-safe), never BMP.
      return { ext: 'png', contentType: 'image/png' };
  }
}

/** P0-1: Detect and validate image format from decoded magic bytes */
function detectImageFormatFromMagic(buf: Buffer): { ext: string; contentType: string } | null {
  if (buf.length >= 4 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) {
    return { ext: 'png', contentType: 'image/png' };
  }
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) {
    return { ext: 'jpeg', contentType: 'image/jpeg' };
  }
  if (buf.length >= 4 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x38) {
    return { ext: 'gif', contentType: 'image/gif' };
  }
  if (
    buf.length >= 12 &&
    buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46 &&
    buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50
  ) {
    return { ext: 'webp', contentType: 'image/webp' };
  }
  if (buf.length >= 2 && buf[0] === 0x42 && buf[1] === 0x4d) {
    return { ext: 'bmp', contentType: 'image/bmp' };
  }
  // TIFF: little-endian II*\0 or big-endian MM\0* (Word-accepted raster).
  if (
    buf.length >= 4 &&
    ((buf[0] === 0x49 && buf[1] === 0x49 && buf[2] === 0x2a && buf[3] === 0x00) ||
      (buf[0] === 0x4d && buf[1] === 0x4d && buf[2] === 0x00 && buf[3] === 0x2a))
  ) {
    return { ext: 'tiff', contentType: 'image/tiff' };
  }
  // SVG: leading BOM/whitespace then <svg, or an XML prolog opening onto <svg.
  // XXE-gated at the call site (SVG is XML — <!DOCTYPE/<!ENTITY still throws).
  if (isSvgContent(buf)) {
    return { ext: 'svg', contentType: 'image/svg+xml' };
  }
  return null;
}

/** True when the buffer opens like an SVG document (prolog-tolerant). */
function isSvgContent(buf: Buffer): boolean {
  const head = buf.subarray(0, 1024).toString('latin1').replace(/^\uFEFF/, '').trimStart();
  if (/^<svg[\s>/]/i.test(head)) return true;
  if (/^<\?xml[\s?]/i.test(head) && /<svg[\s>/]/i.test(head.slice(0, 1024))) return true;
  return false;
}

/**
 * Shared bytes-to-format decode for both image insert shapes: inline
 * `fileData` payloads and jailed `filePath` reads land here with raw bytes
 * and proceed identically. Magic-gated (PNG, JPEG, GIF, WebP, BMP, TIFF,
 * SVG); SVG stays XXE-gated. The default remedy preserves the legacy inline
 * message byte-identical; file reads pass the file-oriented remedy.
 */
function decodeImageBytes(
  imgBuffer: Buffer,
  invalidRemedy = 'Provide image as data: URL or base64 string'
): { ext: string; contentType: string } {
  const magic = detectImageFormatFromMagic(imgBuffer);
  if (!magic) {
    throw toError(
      'E_VALIDATION',
      'Invalid image payload or magic bytes mismatch.',
      invalidRemedy
    );
  }
  if (magic.ext === 'svg') {
    assertNoXxe(imgBuffer.toString('utf8'), 'image SVG payload');
  }
  return { ext: magic.ext, contentType: magic.contentType };
}

// A-1: builds a schema-valid wp:wrap* element for the given mode.
// wrapTight/wrapThrough require wrapText + a wp:wrapPolygon child
// (ECMA-376 CT_WrapTight/CT_WrapThrough); without them the file fails validation.
function buildAnchorWrapElement(doc: Document, wrapMode: ImageWrapMode): Element {
  let wrapTag = 'wp:wrapSquare';
  if (wrapMode === 'tight') wrapTag = 'wp:wrapTight';
  else if (wrapMode === 'through') wrapTag = 'wp:wrapThrough';
  else if (wrapMode === 'topAndBottom') wrapTag = 'wp:wrapTopAndBottom';
  else if (wrapMode === 'behindText' || wrapMode === 'inFrontOfText') wrapTag = 'wp:wrapNone';

  const wrapElement = createWordprocessingDrawingElement(doc, wrapTag);
  if (wrapTag === 'wp:wrapSquare' || wrapTag === 'wp:wrapTight' || wrapTag === 'wp:wrapThrough') {
    wrapElement.setAttribute('wrapText', 'bothSides');
  }
  if (wrapTag === 'wp:wrapTight' || wrapTag === 'wp:wrapThrough') {
    const polygon = createWordprocessingDrawingElement(doc, 'wp:wrapPolygon');
    polygon.setAttribute('edited', '0');
    const start = createWordprocessingDrawingElement(doc, 'wp:start');
    start.setAttribute('x', '0');
    start.setAttribute('y', '0');
    polygon.appendChild(start);
    const corners: Array<[string, string]> = [['0', '21600'], ['21600', '21600'], ['21600', '0'], ['0', '0']];
    for (const [x, y] of corners) {
      const lineTo = createWordprocessingDrawingElement(doc, 'wp:lineTo');
      lineTo.setAttribute('x', x);
      lineTo.setAttribute('y', y);
      polygon.appendChild(lineTo);
    }
    wrapElement.appendChild(polygon);
  }
  return wrapElement;
}

/** True when a paragraph carries an embedded picture (a drawing with an a:blip) next to its text. */
function paragraphHoldsPicture(node: Node): boolean {
  const xml = node.xml as Element | undefined;
  if (!xml || xml.tagName !== 'w:p') return false;
  return xml.getElementsByTagName('w:drawing').length > 0 && xml.getElementsByTagName('a:blip').length > 0;
}

/**
 * Finds the target of an image mutation. Refuses before any change is made:
 * E_NOT_FOUND when the id does not exist, E_VALIDATION when the node is not an image.
 */
function resolveImageTarget(
  ast: DocumentAST,
  targetId: string,
  opName: string
): { node: Node; section: DocumentAST['sections'][number]; index: number; mixed: boolean } {
  for (const section of ast.sections) {
    const index = section.nodes.findIndex((n) => n.id === targetId);
    if (index === -1) continue;
    const node = section.nodes[index];
    const mixed = node.type === 'paragraph' && paragraphHoldsPicture(node);
    if (node.type !== 'image' && !mixed) {
      throw toError(
        'E_VALIDATION',
        `Target '${targetId}' is a ${node.type}, but image.${opName} requires an image node or a paragraph that holds a picture.`,
        'Pass the id of an image node, or of a paragraph that contains a picture (see inspect_document, flag [IMG]).'
      );
    }
    return { node, section, index, mixed };
  }
  const detail = opName === 'setCaption' ? `: Target node '${targetId}' not found for setCaption` : '';
  throw toError(
    'E_NOT_FOUND',
    `Target image node '${targetId}' not found${detail}`,
    're-read the current structure with inspect_document and retry with a valid image node id'
  );
}

export function executeImage(op: ImageOperation, ast: DocumentAST, parser?: DocxParser, allowedDir?: string): { warnings: string[] } | void {
  const relsManipulator = new RelationshipsManipulator(ast);

  // 1. DELETE Operation
  if (op.operation === 'delete') {
    if (!op.target) {
      throw toError(
        'E_VALIDATION',
        'Delete image requires a target node id.',
        "pass { action:'image', operation:'delete', target:'<imageNodeId>' } from inspect_document."
      );
    }
    const { node, section, index, mixed } = resolveImageTarget(ast, op.target, 'delete');
    validateVersion(node, op.expectedVersion);
    if (mixed) {
      for (const drawing of Array.from((node.xml as Element).getElementsByTagName('w:drawing'))) {
        if (drawing.getElementsByTagName('a:blip').length === 0) continue;
        if (drawing.getElementsByTagName('w:txbxContent').length > 0) continue;
        // Remove the drawing alone (and its mc:AlternateContent wrapper), never the whole run:
        // a run can also hold the paragraph text.
        const choice = drawing.parentNode as Element | null;
        const doomed = choice?.nodeName === 'mc:Choice' && choice.parentNode?.nodeName === 'mc:AlternateContent'
          ? (choice.parentNode as Element)
          : drawing;
        const run = doomed.parentNode as Element | null;
        run?.removeChild(doomed);
        const runIsEmpty = run?.nodeName === 'w:r'
          && !Array.from(run.childNodes).some((c) => c.nodeType === 1 && c.nodeName !== 'w:rPr');
        if (run && runIsEmpty) run.parentNode?.removeChild(run);
      }
      node.metadata.hasImage = (node.xml as Element).getElementsByTagName('w:drawing').length > 0;
      node.version += 1;
      section.version += 1;
      return;
    }
    node.xml?.parentNode?.removeChild(node.xml);
    section.nodes.splice(index, 1);
    section.version += 1;
    parser?.removeFromNodeMap(op.target);
    return;
  }

  // 2. RESIZE Operation
  if (op.operation === 'resize') {
    if (!op.target) {
      throw toError('E_VALIDATION', 'Resize image requires a target node id', "pass { action:'image', operation:'resize', target:'<imageNodeId>', width, height }");
    }
    if (op.width === undefined && op.height === undefined) {
      throw toError('E_VALIDATION', 'image.resize requires width and/or height (EMU)', 'pass width, height, or both; with only one, the other follows the current aspect ratio');
    }
    const { node, section } = resolveImageTarget(ast, op.target, 'resize');
    validateVersion(node, op.expectedVersion);
    const xmlNode = node.xml as Element;
    const extents = Array.from(xmlNode.getElementsByTagName('wp:extent')) as Element[];

    let widthEmu = op.width;
    let heightEmu = op.height;
    if (widthEmu === undefined || heightEmu === undefined) {
      const curW = Number(extents[0]?.getAttribute('cx'));
      const curH = Number(extents[0]?.getAttribute('cy'));
      if (!Number.isInteger(curW) || !Number.isInteger(curH) || curW <= 0 || curH <= 0) {
        throw toError('E_VALIDATION', 'Cannot infer aspect ratio: the current image size is unreadable', 'pass both width and height');
      }
      if (widthEmu === undefined) widthEmu = Math.round(((heightEmu as number) * curW) / curH);
      else heightEmu = Math.round((widthEmu * curH) / curW);
    }

    for (const ext of extents) {
      ext.setAttribute('cx', String(widthEmu));
      ext.setAttribute('cy', String(heightEmu));
    }
    const aExts = xmlNode.getElementsByTagName('a:ext');
    for (let i = 0; i < aExts.length; i++) {
      const ext = aExts[i] as Element;
      if (ext.hasAttribute('cx') && ext.hasAttribute('cy')) {
        ext.setAttribute('cx', String(widthEmu));
        ext.setAttribute('cy', String(heightEmu));
      }
    }
    node.version += 1;
    section.version += 1;
    return;
  }

  // 3. SET ALT TEXT / CAPTION Operation
  if (op.operation === 'setAltText') {
    if (!op.target) {
      throw toError(
        'E_VALIDATION',
        'SetAltText requires a target node id',
        "pass { action:'image', operation:'setAltText', target:'<imageNodeId>', altText:'...' }"
      );
    }
    const { node, section } = resolveImageTarget(ast, op.target, 'setAltText');
    validateVersion(node, op.expectedVersion);
    const xmlNode = node.xml as Element;
    const docPrs = xmlNode.getElementsByTagName('wp:docPr');
    if (docPrs.length > 0) {
      const docPr = docPrs[0] as Element;
      if (op.altText) {
        docPr.setAttribute('descr', op.altText);
        docPr.setAttribute('title', op.altText);
        node.version += 1;
        section.version += 1;
      }
    }
    return;
  }

  // 4. SET CAPTION Operation
  if (op.operation === 'setCaption') {
    if (!op.target) {
      throw toError(
        'E_VALIDATION',
        'SetCaption requires a target node id',
        "pass { action:'image', operation:'setCaption', target:'<imageNodeId>', captionText:'...' }"
      );
    }
    if (op.captionText === undefined || op.captionText === null) {
      throw toError(
        'E_VALIDATION',
        'SetCaption requires captionText',
        "pass { action:'image', operation:'setCaption', target:'<imageNodeId>', captionText:'...' }"
      );
    }

    const { node: targetNode, section: targetSection, index: targetIndex } = resolveImageTarget(ast, op.target, 'setCaption');
    validateVersion(targetNode, op.expectedVersion);

    // Check if the target node's next sibling in AST and DOM is already a caption paragraph
    const nextNode = targetSection.nodes[targetIndex + 1];
    let isCaptionSibling = false;
    if (nextNode && nextNode.type === 'paragraph' && nextNode.xml) {
      const pPr = nextNode.xml.getElementsByTagName('w:pPr')[0];
      if (pPr) {
        const pStyle = pPr.getElementsByTagName('w:pStyle')[0];
        if (pStyle && (pStyle.getAttribute('w:val') === 'Caption' || pStyle.getAttribute('val') === 'Caption')) {
          isCaptionSibling = true;
        }
      }
    }

    if (isCaptionSibling && nextNode) {
      // Update existing caption run: clear trailing runs and consolidate into first run
      const runs = Array.from(nextNode.xml.getElementsByTagName('w:r'));
      if (runs.length > 0) {
        for (let i = 1; i < runs.length; i++) {
          runs[i].parentNode?.removeChild(runs[i]);
        }
        let t = runs[0].getElementsByTagName('w:t')[0];
        if (!t) {
          t = createWordElement(nextNode.xml.ownerDocument, 'w:t');
          runs[0].appendChild(t);
        }
        t.textContent = op.captionText;
      } else {
        const doc = nextNode.xml.ownerDocument;
        let r = nextNode.xml.getElementsByTagName('w:r')[0];
        if (!r) {
          r = createWordElement(doc, 'w:r');
          nextNode.xml.appendChild(r);
        }
        const t = createWordElement(doc, 'w:t');
        t.textContent = op.captionText;
        r.appendChild(t);
      }
      nextNode.metadata = {
        ...nextNode.metadata,
        text: op.captionText,
        namedStyle: 'Caption',
      };
      nextNode.version += 1;
      const xmlStr = new XMLSerializer().serializeToString(nextNode.xml);
      nextNode.checksum = crypto.createHash('sha256').update(xmlStr).digest('hex');
    } else {
      // Create fresh caption paragraph element
      const doc = targetNode.xml?.ownerDocument || targetSection.nodes[0]?.xml?.ownerDocument;
      if (!doc) {
        throw toError(
          'E_STATE',
          'Cannot resolve owner XML document for caption',
          're-read the document structure with inspect_document and retry'
        );
      }

      if (!targetNode.xml || !targetNode.xml.parentNode) {
        throw toError(
          'E_STATE',
          `Cannot insert caption: target node '${op.target}' is detached from XML DOM parent.`,
          're-read the document structure with inspect_document and retry'
        );
      }

      const p = createWordElement(doc, 'w:p');
      const captionId = (op as { id?: string }).id || (isImageReproducible()
        ? (() => {
            const base = `caption_${crypto.createHash('sha256').update(`${op.target}:${op.captionText}`).digest('hex').slice(0, 8)}`;
            const live = collectLiveIds(ast);
            let candidate = base;
            let n = 2;
            while (live.has(candidate)) candidate = `${base}_${n++}`;
            return candidate;
          })()
        : `caption_${Date.now()}_${++nodeIdSeq}`);
      p.setAttribute('data-astra-id', captionId);

      const pPr = createWordElement(doc, 'w:pPr');
      const pStyle = createWordElement(doc, 'w:pStyle');
      pStyle.setAttribute('w:val', 'Caption');
      pPr.appendChild(pStyle);
      p.appendChild(pPr);

      const r = createWordElement(doc, 'w:r');
      const t = createWordElement(doc, 'w:t');
      t.textContent = op.captionText;
      r.appendChild(t);
      p.appendChild(r);

      targetNode.xml.parentNode.insertBefore(p, targetNode.xml.nextSibling);

      if (parser) {
        parser.nodeMap?.set?.(captionId, p);
      }

      const xmlStr = new XMLSerializer().serializeToString(p);
      const checksum = crypto.createHash('sha256').update(xmlStr).digest('hex');

      const captionNode: Node = {
        id: captionId,
        type: 'paragraph',
        xml: p,
        metadata: {
          text: op.captionText,
          namedStyle: 'Caption',
          hasImage: false,
          hasCitation: false,
          hasComment: false,
          hasBookmark: false,
          hasTrackedChanges: false,
          hasEquation: false,
          hasField: false,
          hasHyperlink: false,
        },
        version: 1,
        checksum,
      };

      targetSection.nodes.splice(targetIndex + 1, 0, captionNode);
    }

    targetNode.version += 1;
    targetSection.version += 1;
    return;
  }

  // 5. REPOSITION / SET WRAP Operation
  if (op.operation === 'reposition' || op.operation === 'setWrap') {
    if (!op.target) {
      throw toError(
        'E_VALIDATION',
        'Reposition/SetWrap requires a target node id',
        "pass { action:'image', operation:'setWrap', target:'<imageNodeId>', wrapMode:'...' }"
      );
    }
    const { node, section } = resolveImageTarget(ast, op.target, op.operation);
    validateVersion(node, op.expectedVersion);
    const xmlNode = node.xml as Element;

    const drawing = xmlNode.getElementsByTagName('w:drawing')[0] || xmlNode;
    const inlineEl = drawing.getElementsByTagName('wp:inline')[0];
    const anchorEl = drawing.getElementsByTagName('wp:anchor')[0];
    const currentContainer = inlineEl || anchorEl;

    const wrapMode = op.wrapMode || 'square';
    if (wrapMode === 'inline' && anchorEl) {
      // Convert anchor -> inline
      const doc = currentContainer.ownerDocument;
      const newInline = createWordprocessingDrawingElement(doc, 'wp:inline');
      newInline.setAttribute('distT', '0');
      newInline.setAttribute('distB', '0');
      newInline.setAttribute('distL', '0');
      newInline.setAttribute('distR', '0');
      while (anchorEl.firstChild) {
        const child = anchorEl.firstChild;
        if (child.nodeName.startsWith('wp:wrap') || child.nodeName === 'wp:positionH' || child.nodeName === 'wp:positionV' || child.nodeName === 'wp:simplePos') {
          anchorEl.removeChild(child);
        } else {
          newInline.appendChild(child);
        }
      }
      drawing.replaceChild(newInline, anchorEl);
    } else if (wrapMode !== 'inline' && inlineEl) {
      // Convert inline -> anchor
      const doc = currentContainer.ownerDocument;
      const newAnchor = createWordprocessingDrawingElement(doc, 'wp:anchor');
      newAnchor.setAttribute('distT', '0');
      newAnchor.setAttribute('distB', '0');
      newAnchor.setAttribute('distL', '114300');
      newAnchor.setAttribute('distR', '114300');
      newAnchor.setAttribute('simplePos', '0');
      newAnchor.setAttribute('relativeHeight', '251658240');
      newAnchor.setAttribute('behindDoc', wrapMode === 'behindText' ? '1' : '0');
      newAnchor.setAttribute('locked', '0');
      newAnchor.setAttribute('layoutInCell', '1');
      newAnchor.setAttribute('allowOverlap', '1');

      // Wrap XML element (A-1: tight/through carry wrapText + default polygon)
      const wrapElement = buildAnchorWrapElement(doc, wrapMode);

      // OP-07: wp:anchor requires simplePos + positionH/V BEFORE extent/wrap/graphic,
      // and the wrap element must precede a:graphic in schema order. Create fresh
      // positioning nodes first, then re-append content with wrap before graphic.
      const simplePos = createWordprocessingDrawingElement(doc, 'wp:simplePos');
      simplePos.setAttribute('x', '0');
      simplePos.setAttribute('y', '0');
      newAnchor.appendChild(simplePos);

      const positionH = createWordprocessingDrawingElement(doc, 'wp:positionH');
      positionH.setAttribute('relativeFrom', 'column');
      const posHOffset = createWordprocessingDrawingElement(doc, 'wp:posOffset');
      posHOffset.appendChild(doc.createTextNode('0'));
      positionH.appendChild(posHOffset);
      newAnchor.appendChild(positionH);

      const positionV = createWordprocessingDrawingElement(doc, 'wp:positionV');
      positionV.setAttribute('relativeFrom', 'paragraph');
      const posVOffset = createWordprocessingDrawingElement(doc, 'wp:posOffset');
      posVOffset.appendChild(doc.createTextNode('0'));
      positionV.appendChild(posVOffset);
      newAnchor.appendChild(positionV);

      // Partition inline children into ECMA-376 CT_Anchor sequence:
      // simplePos, positionH, positionV, extent, effectExtent?, wrap,
      // docPr, cNvGraphicFramePr?, graphic.
      // simplePos/positionH/positionV are freshly created above; extent and
      // effectExtent must precede wrap, while docPr/cNv follow it.
      const extentNodes: globalThis.Node[] = [];
      const docPrNodes: globalThis.Node[] = [];
      const otherPre: globalThis.Node[] = [];
      const graphicNodes: globalThis.Node[] = [];
      while (inlineEl.firstChild) {
        const child = inlineEl.firstChild;
        inlineEl.removeChild(child);
        const tag = (child as Element).nodeName || '';
        const local = tag.includes(':') ? tag.split(':').pop()! : tag;
        if (local === 'graphic') {
          graphicNodes.push(child);
        } else if (tag.startsWith('wp:wrap') || local === 'positionH' || local === 'positionV' || local === 'simplePos') {
          // Drop stale wrap/positioning carried over from the source; the anchor owns fresh ones.
          continue;
        } else if (local === 'extent' || local === 'effectExtent') {
          extentNodes.push(child);
        } else if (local === 'docPr' || local === 'cNvGraphicFramePr') {
          docPrNodes.push(child);
        } else {
          otherPre.push(child);
        }
      }
      for (const n of extentNodes) newAnchor.appendChild(n);
      for (const n of otherPre) newAnchor.appendChild(n);
      newAnchor.appendChild(wrapElement);
      for (const n of docPrNodes) newAnchor.appendChild(n);
      for (const n of graphicNodes) newAnchor.appendChild(n);
      drawing.replaceChild(newAnchor, inlineEl);
    } else if (wrapMode !== 'inline' && anchorEl) {
      // A-1: re-target an already-anchored drawing. Previously a silent no-op;
      // swap the wrap child in place (schema position: after extent/effectExtent,
      // before docPr/graphic) with a valid element and refresh behindDoc.
      const doc = currentContainer.ownerDocument;
      anchorEl.setAttribute('behindDoc', wrapMode === 'behindText' ? '1' : '0');
      const freshWrap = buildAnchorWrapElement(doc, wrapMode);
      let oldWrap: globalThis.Node | null = null;
      let insertBefore: globalThis.Node | null = null;
      for (const k of Array.from(anchorEl.childNodes)) {
        const tag = (k as Element).nodeName || '';
        const local = tag.includes(':') ? tag.split(':').pop()! : tag;
        if (!oldWrap && tag.startsWith('wp:wrap')) {
          oldWrap = k;
          continue;
        }
        if (!insertBefore && (local === 'docPr' || local === 'cNvGraphicFramePr' || local === 'graphic')) {
          insertBefore = k;
        }
      }
      if (oldWrap) anchorEl.removeChild(oldWrap);
      if (insertBefore) anchorEl.insertBefore(freshWrap, insertBefore);
      else anchorEl.appendChild(freshWrap);
    }
    node.version += 1;
    section.version += 1;
    return;
  }

  // 5. INSERT Operation
  if (op.operation !== 'insert') {
    throw toError(
      'E_VALIDATION',
      `image.${op.operation} is not supported yet`,
      'supported operations: insert, delete, resize, setAltText, setCaption, reposition, setWrap'
    );
  }

  if (!op.anchor) {
    throw toError(
      'E_VALIDATION',
      'Insert image requires an anchor',
      'pass anchor with a valid paragraph or block node id'
    );
  }
  const hasFilePath = typeof op.filePath === 'string' && op.filePath.length > 0;
  if (hasFilePath && op.fileData) {
    throw toError(
      'E_VALIDATION',
      'image insert carries both fileData and filePath (mutually exclusive)',
      'pass exactly one image source: fileData with inline base64 bytes, or filePath with a file inside allowedDir'
    );
  }
  if (!hasFilePath && !op.fileData) throw toError(
    'E_VALIDATION',
    'image insert needs an image source (neither fileData nor filePath present)',
    'pass fileData with inline base64 image bytes, or filePath with a file inside allowedDir (batch --allowed-dir / execute_batch allowedDir)'
  );
  if (parser && !parser.getNode(op.anchor)) {
    throw toError(
      'E_NOT_FOUND',
      `Anchor node '${op.anchor}' not found in document`,
      'Re-read the current structure with inspect_document and retry with a valid node id.'
    );
  }

  // P0-1: Image Path Traversal Prevention
  // Inline payloads accept only data URLs (data:image/...;base64,...) or base64
  // strings — never bare filesystem paths (those arrive as filePath, jailed).
  // Validate magic bytes on decoded imgBuffer (PNG, JPEG, GIF, WebP, BMP, TIFF, SVG).
  // SVG text payloads stay XXE-gated (fail-closed on <!DOCTYPE/<!ENTITY).
  // If an invalid payload or path string is passed, fail closed with E_VALIDATION and remedy: "Provide image as data: URL or base64 string".
  // Batch file references arrive as filePath (jailed read above) and join the
  // shared decode path below, so both input shapes produce identical bytes.
  const trimmed = hasFilePath ? '' : op.fileData!.trim();
  const BASE64_CHARS_RE = /^[A-Za-z0-9+/=\s]+$/;

  let imgBuffer: Buffer;
  let ext: string;
  let contentType: string;

  if (hasFilePath) {
    imgBuffer = readImageFileBytes(op.filePath as string, allowedDir);
    const decoded = decodeImageBytes(imgBuffer, 'use a real png, jpg, gif, webp, bmp, tiff, or svg file inside the allowed directory');
    ext = decoded.ext;
    contentType = decoded.contentType;
  } else if (/^data:/i.test(trimmed)) {
    const dataUrlMatch = DATA_URL_PATTERN.exec(trimmed);
    if (!dataUrlMatch) {
      throw toError(
        'E_VALIDATION',
        'Invalid image data URL format.',
        'Provide image as data: URL or base64 string'
      );
    }
    const rawBase64 = trimmed.slice(dataUrlMatch[0].length);
    if (!BASE64_CHARS_RE.test(rawBase64) || rawBase64.trim().length === 0) {
      throw toError(
        'E_VALIDATION',
        'Invalid base64 in image data URL.',
        'Provide image as data: URL or base64 string'
      );
    }
    imgBuffer = Buffer.from(rawBase64, 'base64');
    const magic = detectImageFormatFromMagic(imgBuffer);
    if (!magic) {
      throw toError(
        'E_VALIDATION',
        'Invalid image payload or magic bytes mismatch.',
        'Provide image as data: URL or base64 string'
      );
    }
    if (magic.ext === 'svg') {
      // SVG is XML: accept the content but keep the XXE fail-closed gate.
      assertNoXxe(imgBuffer.toString('utf8'), 'image SVG payload');
    }
    const fmt = deriveImageFormat(dataUrlMatch[1]);
    ext = magic.ext ?? fmt.ext;
    contentType = magic.contentType ?? fmt.contentType;
  } else {
    // Only accept base64 payloads
    if (!BASE64_CHARS_RE.test(trimmed) || trimmed.length === 0) {
      throw toError(
        'E_VALIDATION',
        'Invalid image data string.',
        'Provide image as data: URL or base64 string'
      );
    }
    imgBuffer = Buffer.from(trimmed, 'base64');
    const decoded = decodeImageBytes(imgBuffer);
    ext = decoded.ext;
    contentType = decoded.contentType;
  }

  const px = decodeImagePixels(imgBuffer, ext);
  const intrinsic = fitImageEmus(px);
  const given = (v: unknown): number | undefined => (typeof v === 'number' && v > 0 ? v : undefined);
  const wGiven = given(op.width);
  const hGiven = given(op.height);
  for (const [label, v] of [['width', wGiven], ['height', hGiven]] as const) {
    if (v !== undefined && v < 9525) {
      throw toError('E_VALIDATION', `image insert ${label} ${v} EMU is smaller than one pixel (9525 EMU)`, 'sizes are integers in EMU (914400 per inch, 12700 per point, 9525 per pixel at 96 dpi); pass for example width:2743200 for 3 inches');
    }
  }
  let widthEmu: number;
  let heightEmu: number;
  if (wGiven !== undefined && hGiven !== undefined) { widthEmu = wGiven; heightEmu = hGiven; }
  else if (wGiven !== undefined) { widthEmu = wGiven; heightEmu = px ? Math.max(1, Math.round((wGiven * px.heightPx) / px.widthPx)) : 3200400; }
  else if (hGiven !== undefined) { heightEmu = hGiven; widthEmu = px ? Math.max(1, Math.round((hGiven * px.widthPx) / px.heightPx)) : 5486400; }
  else { widthEmu = intrinsic?.width ?? 5486400; heightEmu = intrinsic?.height ?? 3200400; }

  // Add to media map and AST relationships under a collision-safe rId.
  // Deterministic default: content-seeded fileName (AST-scoped _2.. loop on
  // collision); wall-clock only when opted out.
  const widthEmuForName = widthEmu;
  const heightEmuForName = heightEmu;
  let fileName: string;
  if (isImageReproducible()) {
    const base = `image_${imageHash8(imgBuffer, ext, widthEmuForName, heightEmuForName)}.${ext}`;
    const taken = new Set<string>(Object.keys(ast.mediaFiles ?? {}));
    fileName = base;
    let n = 2;
    while (taken.has(fileName)) {
      fileName = base.replace(`.${ext}`, `_${n++}.${ext}`);
    }
  } else {
    fileName = `image_${Date.now()}_${++fileNameSeq}.${ext}`;
  }
  const rId = relsManipulator.addMedia(fileName, imgBuffer);

  if (parser) {
    const zipAdapter = parser.getZipAdapter();
    const domParser = new DOMParser();
    const serializer = new XMLSerializer();

    // Write media file directly to zip
    zipAdapter.writeBuffer(`word/media/${fileName}`, imgBuffer);

    const relsPath = 'word/_rels/document.xml.rels';
    const relsXml = zipAdapter.readTextOrNull(relsPath);
    if (relsXml) assertNoXxe(relsXml, relsPath);
    let relsDoc = relsXml ? domParser.parseFromString(relsXml, 'text/xml') : null;
    if (!relsDoc || !relsDoc.documentElement || relsDoc.getElementsByTagName('parsererror').length > 0) {
      const freshRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
      relsDoc = domParser.parseFromString(freshRels, 'text/xml');
    }
    const relsRoot = relsDoc.documentElement;
    const newRel = createRelationshipsElement(relsDoc, 'Relationship');
    newRel.setAttribute('Id', rId);
    newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
    newRel.setAttribute('Target', `media/${fileName}`);
    relsRoot.appendChild(newRel);
    zipAdapter.writeText(relsPath, serializer.serializeToString(relsDoc));

    // Write Content Types default for the derived extension
    const typesPath = '[Content_Types].xml';
    const typesXml = zipAdapter.readTextOrNull(typesPath);
    if (typesXml) {
      assertNoXxe(typesXml, typesPath);
      const typesDoc = domParser.parseFromString(typesXml, 'text/xml');
      const typesRoot = typesDoc.documentElement;
      const defaults = Array.from(typesRoot.getElementsByTagName('Default'));
      const alreadyDeclared = defaults.some((def) => (def as Element).getAttribute('Extension') === ext);
      if (!alreadyDeclared) {
        const newDef = createContentTypesElement(typesDoc, 'Default');
        newDef.setAttribute('Extension', ext);
        newDef.setAttribute('ContentType', contentType);
        typesRoot.appendChild(newDef);
      }
      zipAdapter.writeText(typesPath, serializer.serializeToString(typesDoc));
    }
  }

  // Create drawing XML for the image
  const safeName = isImageReproducible()
    ? `Image_${imageHash8(imgBuffer, ext, widthEmu, heightEmu)}`
    : `Image_${Date.now()}`;
  // A-2: scan the live DOM for existing wp:docPr/pic:cNvPr ids, allocate max+1.
  let drawingDom: Document | null = null;
  try {
    drawingDom = parser ? parser.getDocumentDom() : null;
  } catch { drawingDom = null; }
  const docPrId = allocateNextDrawingDocPrId(drawingDom);

  // Alt text is preserved on the drawing (descr + title) so exports render
  // the author's text. Missing alt warns (never a silent node-id alt).
  const escapeAltAttr = (s: string): string => s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
  const altProvided = typeof op.altText === 'string' && op.altText.trim().length > 0;
  const altClean = altProvided ? (op.altText as string).trim() : '';
  const docPrAltAttrs = altProvided
    ? ` descr="${escapeAltAttr(altClean)}" title="${escapeAltAttr(altClean)}"`
    : '';

  const drawingXml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:pPr><w:spacing w:before="140" w:after="140" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
            <wp:docPr id="${docPrId}" name="${safeName}"${docPrAltAttrs}/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="${docPrId}" name="${safeName}"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${rId}"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${widthEmu}" cy="${heightEmu}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>`;

  const imgId = (op as { id?: string }).id || (isImageReproducible()
    ? (() => {
        const base = `img_${imageHash8(imgBuffer, ext, widthEmu, heightEmu)}`;
        const live = collectLiveIds(ast);
        try {
          if (parser) {
            const dom = parser.getDocumentDom();
            const all = dom.getElementsByTagName('*');
            for (let i = 0; i < all.length; i++) {
              const v = (all[i] as Element).getAttribute?.('data-astra-id');
              if (v) live.add(v);
            }
          }
        } catch { /* best-effort */ }
        let candidate = base;
        let n = 2;
        while (live.has(candidate)) candidate = `${base}_${n++}`;
        return candidate;
      })()
    : `img_${Date.now()}_${++nodeIdSeq}`);
  let drawingNodeId: string | undefined;
  if (parser) {
    drawingNodeId = parser.insertRootNode(drawingXml, op.anchor, 'after', imgId);
  }

  if (ast.sections.length > 0) {
    let imageXmlNode: Element | undefined;
    if (parser && drawingNodeId) {
      imageXmlNode = parser.getNode(drawingNodeId);
    }
    if (!imageXmlNode) {
      const doc = new DOMParser().parseFromString(drawingXml, 'text/xml');
      imageXmlNode = doc.documentElement;
    }

    const newNode: Node = {
      id: imgId,
      type: 'image',
      xml: imageXmlNode as Element,
      metadata: {
        text: '',
        hasImage: true,
        hasCitation: false,
        hasComment: false,
        hasBookmark: false,
        hasTrackedChanges: false,
        hasEquation: false,
        hasField: false,
        hasHyperlink: false
      },
      version: 1,
      checksum: `img_chk_${rId}_${imgBuffer.length}`
    };

    // OP-06: splice into the anchor's own section, not always sections[0].
    // With no anchor the image goes to the end of the document (the DOM
    // side appends before the body's sectPr), so the last section.
    const placeAfter = op.anchor ? locateNode(ast.sections, op.anchor) : undefined;
    if (placeAfter) {
      insertIntoSection(ast.sections, placeAfter.section, placeAfter.index + 1, [newNode]).version += 1;
    } else {
      const endSection = lastSection(ast.sections);
      if (endSection) {
        endSection.nodes.push(newNode);
        endSection.version += 1;
      }
    }
  }

  // A-2: honor wrapMode and captionText on image.insert (Step-0 quarantined
  // hunk, IMAGE part only — safe now that A-1 makes tight/through valid).
  if (op.wrapMode && op.wrapMode !== 'inline') {
    executeImage({ action: 'image', operation: 'setWrap', target: imgId, wrapMode: op.wrapMode }, ast, parser);
  }
  if (op.captionText) {
    executeImage({ action: 'image', operation: 'setCaption', target: imgId, captionText: op.captionText }, ast, parser);
  }
  if (!altProvided) {
    return {
      warnings: [
        `Image '${imgId}' has no alt text: exports render it as a generic image. ` +
          `Remedy: pass altText on insert, or set it later with { action:'image', operation:'setAltText', target:'${imgId}', altText:'...' }.`,
      ],
    };
  }
}
