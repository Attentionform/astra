import type { ChartFloatingStyle } from '../operations/chart-style';

/** Builds the inline or floating wp container for a chart drawing. */
export function buildChartDrawingContainer(args: {
  rId: string;
  docPrId: number;
  name: string;
  widthEmu: number;
  heightEmu: number;
  floating?: ChartFloatingStyle;
  /** Alternative text, written to `wp:docPr/@descr` (not to the chart part). */
  altText?: string;
}): string {
  const docPr = `<wp:docPr id="${args.docPrId}" name="${args.name}"${chartAltTextAttr(args.altText)}/>`;
  const graphic = `<a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart r:id="${args.rId}"/>
          </a:graphicData>
        </a:graphic>`;
  if (args.floating === undefined) {
    return `<wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="${args.widthEmu}" cy="${args.heightEmu}"/>
        ${docPr}
        ${graphic}
      </wp:inline>`;
  }
  const horizontal = args.floating.horizontalAlign ?? 'center';
  const vertical = args.floating.verticalAlign === 'middle'
    ? 'center'
    : args.floating.verticalAlign ?? 'top';
  const wrap = args.floating.wrap === 'none'
    ? '<wp:wrapNone/>'
    : '<wp:wrapSquare wrapText="bothSides"/>';
  return `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="${args.floating.behindText ? '1' : '0'}" locked="0" layoutInCell="1" allowOverlap="0">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="margin"><wp:align>${horizontal}</wp:align></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:align>${vertical}</wp:align></wp:positionV>
        <wp:extent cx="${args.widthEmu}" cy="${args.heightEmu}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        ${wrap}
        ${docPr}
        ${graphic}
      </wp:anchor>`;
}

function escapeAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** ` descr="..."` for a non-empty alt text, else an empty string. */
export function chartAltTextAttr(altText: string | undefined): string {
  const trimmed = altText?.trim();
  return trimmed ? ` descr="${escapeAttr(trimmed)}"` : '';
}

interface DrawingLike {
  getElementsByTagName(name: string): ArrayLike<DrawingLike>;
  getAttribute(name: string): string | null;
  setAttribute(name: string, value: string): void;
  removeAttribute(name: string): void;
}

/** The `wp:inline` or `wp:anchor` that carries the chart graphic inside a body node, or null. */
export function findChartDrawingContainer(nodeXml: unknown): DrawingLike | null {
  const root = nodeXml as DrawingLike | undefined;
  if (!root || typeof root.getElementsByTagName !== 'function') return null;
  for (const tag of ['wp:inline', 'wp:anchor']) {
    const containers = root.getElementsByTagName(tag);
    for (let i = 0; i < containers.length; i++) {
      if (containers[i].getElementsByTagName('c:chart').length > 0) return containers[i];
    }
  }
  return null;
}

/** True when the chart drawing is floating (`wp:anchor`), false when inline, null when no drawing was found. */
export function isChartDrawingFloating(nodeXml: unknown): boolean | null {
  const container = findChartDrawingContainer(nodeXml) as (DrawingLike & { tagName?: string; nodeName?: string }) | null;
  if (!container) return null;
  return (container.tagName ?? container.nodeName) === 'wp:anchor';
}

/** Reads the drawing size in EMU, or null when there is no chart drawing. */
export function readChartDrawingSize(nodeXml: unknown): { widthEmu: number; heightEmu: number } | null {
  const container = findChartDrawingContainer(nodeXml);
  const extent = container?.getElementsByTagName('wp:extent')[0];
  if (!extent) return null;
  const widthEmu = Number(extent.getAttribute('cx'));
  const heightEmu = Number(extent.getAttribute('cy'));
  return Number.isFinite(widthEmu) && Number.isFinite(heightEmu) ? { widthEmu, heightEmu } : null;
}

/** Sets the drawing size. A dimension that is not given keeps its current value. Returns false when no drawing exists. */
export function setChartDrawingSize(nodeXml: unknown, size: { widthEmu?: number; heightEmu?: number }): boolean {
  const container = findChartDrawingContainer(nodeXml);
  const extent = container?.getElementsByTagName('wp:extent')[0];
  if (!extent) return false;
  if (size.widthEmu !== undefined) extent.setAttribute('cx', String(size.widthEmu));
  if (size.heightEmu !== undefined) extent.setAttribute('cy', String(size.heightEmu));
  return true;
}

/** Sets (non-empty text) or clears (empty text) the alt text on the chart's `wp:docPr`. Returns false when no drawing exists. */
export function setChartAltText(nodeXml: unknown, altText: string): boolean {
  const container = findChartDrawingContainer(nodeXml);
  const docPr = container?.getElementsByTagName('wp:docPr')[0];
  if (!docPr) return false;
  const trimmed = altText.trim();
  if (trimmed) docPr.setAttribute('descr', trimmed);
  else docPr.removeAttribute('descr');
  return true;
}
