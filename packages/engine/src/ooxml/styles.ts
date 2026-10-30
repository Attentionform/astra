/**
 * ISO/IEC 29500 (ECMA-376) Style Palette Introspection Engine
 * 
 * Extracts style palettes, brand colors, and typography hierarchies from styles.xml
 * so AI agents can introspect available styles instead of hallucinating formatting rules.
 */

import { DOMParser } from '@xmldom/xmldom';
import type { DocxParser } from '../parser';
import { assertNoXxe } from './xxe-guard';

export interface StyleInfo {
  id: string;
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  isDefault?: boolean;
  basedOn?: string;
  next?: string;
  font?: string;
  fontSizePt?: number;
  color?: string; // hex without '#'
  bold?: boolean;
  italic?: boolean;
  quickFormat?: boolean;
}

export interface StylePalette {
  primaryFont: string;
  headingFont?: string;
  brandColors: string[]; // hex codes without '#'
  availableStyles: StyleInfo[];
  headingStyles: string[]; // ['Heading1', 'Heading2', ...]
}

/**
 * Parses word/styles.xml and extracts a rich StylePalette.
 */
export function extractStylePalette(stylesDocOrXml: Document | string): StylePalette {
  if (!stylesDocOrXml || (typeof stylesDocOrXml === 'string' && stylesDocOrXml.trim() === '')) {
    return {
      primaryFont: 'Calibri',
      brandColors: [],
      availableStyles: [],
      headingStyles: []
    };
  }

  let doc: Document;
  if (typeof stylesDocOrXml === 'string') {
    assertNoXxe(stylesDocOrXml, 'word/styles.xml');
    const parser = new DOMParser({
      errorHandler: {
        warning: () => {},
        error: () => {},
        fatalError: (e) => { throw new Error(e); }
      }
    });
    doc = parser.parseFromString(stylesDocOrXml, 'text/xml');
  } else {
    doc = stylesDocOrXml;
  }

  let primaryFont = 'Calibri';
  let headingFont: string | undefined = undefined;
  const brandColorsSet = new Set<string>();
  const availableStyles: StyleInfo[] = [];
  const headingStyles: string[] = [];

  // 1. Extract document defaults from <w:docDefaults>
  const docDefaults = doc.getElementsByTagName('w:docDefaults')[0];
  if (docDefaults) {
    const rFonts = docDefaults.getElementsByTagName('w:rFonts')[0];
    if (rFonts) {
      const ascii = rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi');
      if (ascii) primaryFont = ascii;
    }
  }

  // 2. Iterate through all <w:style> elements
  const styleEls = doc.getElementsByTagName('w:style');
  for (let i = 0; i < styleEls.length; i++) {
    const el = styleEls[i];
    const styleId = el.getAttribute('w:styleId') || '';
    const styleType = (el.getAttribute('w:type') || 'paragraph') as StyleInfo['type'];
    const isDefault = el.getAttribute('w:default') === '1' || el.getAttribute('w:default') === 'true';

    let name = styleId;
    const nameEl = el.getElementsByTagName('w:name')[0];
    if (nameEl && nameEl.getAttribute('w:val')) {
      name = nameEl.getAttribute('w:val')!;
    }

    let basedOn: string | undefined;
    const basedOnEl = el.getElementsByTagName('w:basedOn')[0];
    if (basedOnEl && basedOnEl.getAttribute('w:val')) {
      basedOn = basedOnEl.getAttribute('w:val')!;
    }

    let next: string | undefined;
    const nextEl = el.getElementsByTagName('w:next')[0];
    if (nextEl && nextEl.getAttribute('w:val')) {
      next = nextEl.getAttribute('w:val')!;
    }

    let font: string | undefined;
    let fontSizePt: number | undefined;
    let color: string | undefined;
    let bold: boolean | undefined;
    let italic: boolean | undefined;

    const rPr = el.getElementsByTagName('w:rPr')[0];
    if (rPr) {
      const rFonts = rPr.getElementsByTagName('w:rFonts')[0];
      if (rFonts) {
        font = rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi') || undefined;
      }
      const sz = rPr.getElementsByTagName('w:sz')[0];
      if (sz && sz.getAttribute('w:val')) {
        const halfPoints = parseInt(sz.getAttribute('w:val')!, 10);
        if (!isNaN(halfPoints)) fontSizePt = halfPoints / 2;
      }
      const colorEl = rPr.getElementsByTagName('w:color')[0];
      if (colorEl && colorEl.getAttribute('w:val')) {
        const val = colorEl.getAttribute('w:val')!.trim();
        if (val && val.toLowerCase() !== 'auto') {
          color = val.toUpperCase().replace(/^#/, '');
          brandColorsSet.add(color);
        }
      }
      if (rPr.getElementsByTagName('w:b').length > 0) bold = true;
      if (rPr.getElementsByTagName('w:i').length > 0) italic = true;
    }

    const qFormat = el.getElementsByTagName('w:qFormat').length > 0;

    const info: StyleInfo = {
      id: styleId,
      name,
      type: styleType,
      isDefault: isDefault ? true : undefined,
      basedOn,
      next,
      font,
      fontSizePt,
      color,
      bold,
      italic,
      quickFormat: qFormat ? true : undefined
    };

    availableStyles.push(info);

    if (styleType === 'paragraph') {
      if (/^Heading\d/i.test(styleId) || /^Heading\s*\d/i.test(name)) {
        headingStyles.push(styleId);
        if (!headingFont && font) {
          headingFont = font;
        }
      }
    }
  }

  // Deduplicate and filter brand colors (must be valid 6-char hex, exclude pure black/white)
  const brandColors = Array.from(brandColorsSet).filter(c => /^[0-9A-F]{6}$/i.test(c) && c !== '000000' && c !== 'FFFFFF');

  return {
    primaryFont,
    headingFont,
    brandColors,
    availableStyles,
    headingStyles
  };
}

/**
 * Formats a style palette into a token-efficient summary for LLM context prompts.
 */
export function formatPaletteForPrompt(palette: StylePalette): string {
  const lines: string[] = [];
  lines.push(`Primary Font: ${palette.primaryFont}`);
  if (palette.headingFont) {
    lines.push(`Heading Font: ${palette.headingFont}`);
  }
  if (palette.brandColors.length > 0) {
    lines.push(`Brand Colors: ${palette.brandColors.join(', ')}`);
  }
  if (palette.headingStyles.length > 0) {
    lines.push('Available Heading Styles:');
    for (const h of palette.headingStyles) {
      const s = palette.availableStyles.find(item => item.id === h);
      if (s) {
        const details: string[] = [];
        if (s.font) details.push(`font=${s.font}`);
        if (s.fontSizePt) details.push(`${s.fontSizePt}pt`);
        if (s.color) details.push(`#${s.color}`);
        if (s.bold) details.push('bold');
        lines.push(`  - ${s.id} ("${s.name}"): ${details.join(', ')}`);
      } else {
        lines.push(`  - ${h}`);
      }
    }
  } else {
    lines.push('Available Heading Styles: None');
  }

  const customStyles = palette.availableStyles.filter(s => 
    !palette.headingStyles.includes(s.id) && 
    s.id !== 'Normal' && 
    s.id !== 'DefaultParagraphFont'
  );
  if (customStyles.length > 0) {
    lines.push('Other Document Styles:');
    for (const s of customStyles.slice(0, 15)) {
      lines.push(`  - ${s.id} ("${s.name}", type=${s.type})`);
    }
  }

  return lines.join('\n');
}

/**
 * Helper to extract style palette from a DocxParser instance.
 */
export function getTemplatePalette(parser: DocxParser): StylePalette {
  return parser.getStylePalette();
}
