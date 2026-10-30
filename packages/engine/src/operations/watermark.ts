import { DocumentAST } from '../ast';
import { DocxParser } from '../parser';
import { BaseOperation } from './base';
import { HeaderFooterManipulator } from '../manipulators/header-footer';

export interface WatermarkOperation extends BaseOperation {
  action: 'watermark';
  text?: string;
  color?: string;      // HEX without '#', default 'D1D5DB' (light gray)
  fontSize?: number;   // default 60pt
  opacity?: number;    // default 0.3
  rotation?: number;   // default 315 (-45 deg diagonal)
}

export function executeWatermark(op: WatermarkOperation, ast: DocumentAST, parser?: DocxParser): void {
  const rawText = (op.text || 'CONFIDENTIAL').trim();
  // OP-18: escape for use inside an XML attribute (v:textpath string="...").
  const text = rawText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
  const color = (/^[0-9A-Fa-f]{6}$/.test((op.color || '').replace(/^#/, ''))
    ? (op.color as string).replace(/^#/, '')
    : 'D1D5DB');
  const fontSize = op.fontSize || 60;
  const rotation = op.rotation !== undefined ? op.rotation : 315;

  // Build WordArt/DrawingML XML for Header watermark with behindDoc="1"
  const watermarkXml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:v="urn:schemas-microsoft-com:vml"
                            xmlns:o="urn:schemas-microsoft-com:office:office"
                            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:pPr>
      <w:pStyle w:val="Header"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:noProof/>
      </w:rPr>
      <w:pict>
        <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"
                      path="m@7,l@8,m@5,21600l@6,21600e">
          <v:formulas>
            <v:f eqn="val #0"/>
            <v:f eqn="prod #0 4 3"/>
            <v:f eqn="prod #0 2 3"/>
            <v:f eqn="prod #0 1 3"/>
            <v:f eqn="sum 21600 0 @1"/>
            <v:f eqn="sum 21600 0 @2"/>
            <v:f eqn="sum 21600 0 @3"/>
            <v:f eqn="sum @0 0 10800"/>
            <v:f eqn="prod @7 1 2"/>
          </v:formulas>
          <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@8,0;@5,10800;@8,21600;@6,10800" o:connectangles="270,180,90,0"/>
          <v:textpath on="t" fitshape="t"/>
          <o:lock v:ext="edit" text="t"/>
        </v:shapetype>
        <v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s1025" type="#_x0000_t136"
                 style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:120pt;rotation:${rotation};z-index:-251656704;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
                 o:allowincell="f" fillcolor="#${color}" stroked="f">
          <v:fill opacity="${op.opacity ?? 0.3}"/>
          <v:textpath style="font-family:'Calibri';font-size:${fontSize}pt;font-weight:bold" string="${text}"/>
          <w10:wrap type="none"/>
          <w10:anchorlock/>
        </v:shape>
      </w:pict>
    </w:r>
  </w:p>`;

  if (parser) {
    const manipulator = new HeaderFooterManipulator(ast, parser);
    // OP-18: never wipe existing header content. When a header part already
    // exists, pass the bare watermark paragraph so the manipulator appends it
    // into the live w:hdr root; only synthesize a full w:hdr wrapper when no
    // header exists yet (a bare w:p part without w:hdr root is corrupt).
    const hasHeaderPart = hasExistingHeaderPart(ast, parser);
    const xmlContent = hasHeaderPart
      ? watermarkXml
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
        ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
        ` xmlns:v="urn:schemas-microsoft-com:vml"` +
        ` xmlns:o="urn:schemas-microsoft-com:office:office"` +
        ` xmlns:w10="urn:schemas-microsoft-com:office:word">${watermarkXml}</w:hdr>`;
    manipulator.setHeaderFooter({
      targetType: 'header',
      pageType: 'default',
      xmlContent
    });
  }
}

/** True when the package already carries a header part with live XML. */
function hasExistingHeaderPart(ast: DocumentAST, parser: DocxParser): boolean {
  const headerRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
  try {
    for (const rel of Object.values(ast.relationships || {})) {
      if (rel.type === headerRelType && /(?:^|\/)header\d+\.xml$/.test(rel.target)) {
        const zipPath = rel.target.startsWith('word/') ? rel.target : `word/${rel.target.replace(/^\//, '')}`;
        try {
          if (parser.getZipAdapter().hasFile(zipPath)) return true;
        } catch { /* treat as missing */ }
      }
    }
  } catch { /* fall through */ }
  return false;
}
