import type { ChartFloatingStyle } from '../operations/chart-style';

/** Builds the inline or floating wp container for a chart drawing. */
export function buildChartDrawingContainer(args: {
  rId: string;
  docPrId: number;
  name: string;
  widthEmu: number;
  heightEmu: number;
  floating?: ChartFloatingStyle;
}): string {
  const graphic = `<a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart r:id="${args.rId}"/>
          </a:graphicData>
        </a:graphic>`;
  if (args.floating === undefined) {
    return `<wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="${args.widthEmu}" cy="${args.heightEmu}"/>
        <wp:docPr id="${args.docPrId}" name="${args.name}"/>
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
        <wp:docPr id="${args.docPrId}" name="${args.name}"/>
        ${graphic}
      </wp:anchor>`;
}
