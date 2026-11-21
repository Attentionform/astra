import type { Theme } from './design';

const sizes = { body: 22, title: 56, subtitle: 28, h1: 36, h2: 28, h3: 24, h4: 22, small: 18, code: 20 };

/** Calibri family, deep teal accent. Modern and neutral. */
export const TEAL: Theme = {
  name: 'Teal',
  fonts: { body: 'Calibri', heading: 'Calibri', mono: 'Consolas' },
  colors: { ink: '1F2933', muted: '5B6770', accent: '0B5C6B', accentSoft: 'E6F1F3', rule: 'C5CED3', wash: 'F3F5F6', placeholder: '6B7780' },
  sizes,
};

/** Cambria headings over Calibri body, navy accent. Formal. */
export const NAVY: Theme = {
  name: 'Navy',
  fonts: { body: 'Calibri', heading: 'Cambria', mono: 'Consolas' },
  colors: { ink: '1B2430', muted: '566170', accent: '1F3A5F', accentSoft: 'E8EDF4', rule: 'C9D1DC', wash: 'F2F4F7', placeholder: '666F7D' },
  sizes,
};

/** Cambria body and headings, graphite with a warm brick accent. Editorial. */
export const BRICK: Theme = {
  name: 'Brick',
  fonts: { body: 'Cambria', heading: 'Calibri', mono: 'Consolas' },
  colors: { ink: '262626', muted: '5F5A55', accent: '9A3B26', accentSoft: 'F7ECE8', rule: 'D4CCC6', wash: 'F5F2EF', placeholder: '6F6963' },
  sizes: { ...sizes, body: 22 },
};

export const CANDIDATES: Theme[] = [TEAL, NAVY, BRICK];

/** The design direction chosen in T0 (2026-11-20): Navy. Cambria headings over Calibri body read as formal and neutral, both ship with Word and have metric-compatible fallbacks (Caladea, Carlito) in LibreOffice. */
export const DESIGN: Theme = NAVY;
export const DESIGN_VERSION = 1;

/** Standard manuscript format: Times New Roman 12 pt, double spaced, black on white. */
export const MANUSCRIPT: Theme = {
  name: 'Manuscript',
  variant: 'manuscript',
  fonts: { body: 'Times New Roman', heading: 'Times New Roman', mono: 'Courier New' },
  colors: { ink: '111111', muted: '444444', accent: '111111', accentSoft: 'F2F2F2', rule: 'BBBBBB', wash: 'F5F5F5', placeholder: '595959' },
  sizes: { body: 24, title: 28, subtitle: 24, h1: 24, h2: 24, h3: 24, h4: 24, small: 20, code: 22 },
};
