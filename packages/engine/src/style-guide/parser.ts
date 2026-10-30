/**
 * Line-based `style-guide.md` parser with fail-safe defaulting.
 *
 * Never throws on user input: missing sections and unparseable values fall
 * back to `./defaults` and every assumption lands in `warnings` (section
 * defaults, corrected hex, unknown theme/margin/version, ignored lines —
 * value problems name their 1-based line number).
 *
 * Hex policy: `#` prefix and case are canonical spellings (normalized
 * silently); 3-char shorthand is expanded with a warning; anything else is
 * malformed (default applied + warning).
 */
import { createDefaultStyleGuide } from './defaults';
import {
  STYLE_GUIDE_CALLOUT_THEMES,
  STYLE_GUIDE_MARGINS,
  type StyleGuideCalloutTag,
  type StyleGuideCalloutTheme,
  type StyleGuideConfig,
  type StyleGuideMargin,
  type StyleGuideParseResult,
} from './types';

/** Directory (repo-relative) where bundled guide markdown files live. */
export const BUNDLED_GUIDE_DIR = 'packages/engine/templates/guides';

/**
 * Named-guide path convention: `corporate` ->
 * `packages/engine/templates/guides/corporate.md`. Names are trimmed,
 * lowercased, and a trailing `.md` is accepted. Bundled files themselves
 * ship in a later slice; this helper pins the convention only.
 */
export function bundledGuidePath(name: string): string {
  const base = String(name ?? '').trim().toLowerCase().replace(/\.md$/, '');
  return `${BUNDLED_GUIDE_DIR}/${base}.md`;
}

/** True for file-backed refs (`./brand.md`, `/workspace/g.md`, `C:\g.md`); else a named guide. */
export function isFileGuideRef(ref: string): boolean {
  const v = String(ref ?? '').trim();
  return (
    v.toLowerCase().endsWith('.md') || v.includes('/') || v.includes('\\')
  );
}

/**
 * Normalize a user hex spelling to OOXML 6-char uppercase (no `#`).
 * Returns null when the value is malformed (caller defaults + warns).
 */
export function normalizeHexColor(raw: unknown): string | null {
  if (typeof raw !== 'string') return null;
  const v = raw.trim().replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{6}$/.test(v)) return v;
  return null;
}

/** True only for exact 3-char shorthand (`#RGB`); caller expands + warns. */
export function isShorthandHex(raw: unknown): boolean {
  if (typeof raw !== 'string') return false;
  return /^[0-9A-Fa-f]{3}$/.test(raw.trim().replace(/^#/, ''));
}

export function expandShorthandHex(raw: string): string {
  const v = raw.trim().replace(/^#/, '').toUpperCase();
  return v.split('').map((c) => c + c).join('');
}

function parseBoolean(raw: string): boolean | null {
  const v = raw.trim().toLowerCase();
  if (v === 'true' || v === 'yes' || v === 'on' || v === '1') return true;
  if (v === 'false' || v === 'no' || v === 'off' || v === '0') return false;
  return null;
}

function normKey(raw: string): string {
  return raw.trim().toLowerCase().replace(/[\s\-_]+/g, '');
}

function stripBullet(line: string): string | null {
  const m = /^[\s]*[-*][\s]+(.*)$/.exec(line);
  return m ? m[1].trim() : null;
}

function splitKeyValue(text: string): { key: string; value: string } | null {
  const idx = text.indexOf(':');
  if (idx <= 0) return null;
  return {
    key: text.slice(0, idx).trim(),
    value: text.slice(idx + 1).trim(),
  };
}

const SECTION_ALIASES: Record<string, string> = {
  palette: 'palette',
  typography: 'typography',
  tables: 'tables',
  table: 'tables',
  callouts: 'callouts',
  callout: 'callouts',
  codeblocks: 'codeBlocks',
  codeblock: 'codeBlocks',
  code: 'codeBlocks',
  layout: 'layout',
};

const KNOWN_SECTIONS = [
  'palette',
  'tables',
  'typography',
  'callouts',
  'codeBlocks',
  'layout',
] as const;

function isCalloutTheme(v: string): v is StyleGuideCalloutTheme {
  return (STYLE_GUIDE_CALLOUT_THEMES as readonly string[]).includes(v);
}

function isMargin(v: string): v is StyleGuideMargin {
  return (STYLE_GUIDE_MARGINS as readonly string[]).includes(v);
}

/**
 * Parse guide markdown into `{ config, warnings }`. Total function over
 * strings — any input (empty, garbage, unknown version) yields a valid
 * config; unknown major versions parse best-effort under the v1 schema.
 */
export function parseStyleGuide(markdown: string): StyleGuideParseResult {
  const warnings: string[] = [];
  const config = createDefaultStyleGuide();
  try {
    if (typeof markdown !== 'string') {
      warnings.push('input is not text; full built-in defaults applied');
      return { config, warnings };
    }
    const seenSections = new Set<string>();
    const seenKeys = new Set<string>();
    let current: string | null = null;
    let titleSeen = false;
    const descriptionLines: string[] = [];

    const lines = markdown.split(/\r?\n/);
    let hasDescription = false;
    for (let i = 0; i < lines.length; i++) {
      const lineNo = i + 1;
      let line = lines[i];
      // Strip inline HTML comments; skip comment-only/empty lines.
      line = line.replace(/<!--.*?-->/g, '').trim();
      if (line === '') continue;
      if (/^<!--/.test(lines[i].trim())) continue;

      const h2 = /^##\s+(.*)$/.exec(line);
      if (h2) {
        const name = normKey(h2[1]);
        const mapped = SECTION_ALIASES[name];
        if (mapped) {
          current = mapped;
          seenSections.add(mapped);
        } else {
          warnings.push(
            `line ${lineNo}: unknown section '${h2[1].trim()}' ignored`,
          );
          current = null;
        }
        continue;
      }

      const h1 = /^#\s+(.*)$/.exec(line);
      if (h1 && !titleSeen) {
        titleSeen = true;
        const title = h1[1].trim().replace(/^style guide\s*:\s*/i, '');
        if (title !== '') config.name = title;
        else warnings.push(`line ${lineNo}: empty guide title; default name kept`);
        current = null;
        continue;
      }
      if (h1) {
        warnings.push(`line ${lineNo}: duplicate guide title '${h1[1].trim()}' ignored; kept '${config.name}'`);
        continue;
      }

      const bullet = stripBullet(line);
      if (bullet !== null) {
        const kv = splitKeyValue(bullet);
        if (kv === null) {
          warnings.push(`line ${lineNo}: unparseable bullet ignored ('${bullet}')`);
          continue;
        }
        const key = normKey(kv.key);
        if (key === 'version' && current === null) {
          const n = Number(kv.value);
          if (n === 1) {
            // known version — nothing to do
          } else {
            warnings.push(
              `line ${lineNo}: unknown guide version '${kv.value}'; parsing best-effort under v1 schema`,
            );
          }
          continue;
        }
        if (current === null) {
          warnings.push(`line ${lineNo}: bullet outside any section ignored ('${kv.key}')`);
          continue;
        }
        const dupKey = `${current}:${key}`;
        if (seenKeys.has(dupKey)) {
          warnings.push(
            `line ${lineNo}: duplicate key '${kv.key}' in '${current}' overwrites earlier value; kept '${kv.value}'`,
          );
        } else {
          seenKeys.add(dupKey);
        }
        applyKey(config, warnings, current, kv.key, kv.value, lineNo);
        continue;
      }

      // Bare `Version: N` line outside a section (non-bullet spelling).
      const bare = splitKeyValue(line);
      if (bare && normKey(bare.key) === 'version' && current === null) {
        if (Number(bare.value) !== 1) {
          warnings.push(
            `line ${lineNo}: unknown guide version '${bare.value}'; parsing best-effort under v1 schema`,
          );
        }
        continue;
      }

      if (current === null) {
        if (/^#{1,6}\s/.test(line)) {
          warnings.push(`line ${lineNo}: heading outside any section ignored ('${line}')`);
          continue;
        }
        descriptionLines.push(line);
        hasDescription = true;
        continue;
      }
      warnings.push(`line ${lineNo}: non-bullet line inside '${current}' ignored`);
    }

    if (hasDescription) config.description = descriptionLines.join('\n');
    else delete config.description;

    for (const s of KNOWN_SECTIONS) {
      if (!seenSections.has(s)) {
        warnings.push(`section '${s}' missing; built-in defaults applied`);
      }
    }
    return { config, warnings };
  } catch (err) {
    warnings.push(
      `parser recovered from internal error (${err instanceof Error ? err.message : String(err)}); defaults applied where unset`,
    );
    return { config, warnings };
  }
}

function hexField(
  config: StyleGuideConfig,
  warnings: string[],
  set: (v: string) => void,
  fallback: string,
  label: string,
  value: string,
  lineNo: number,
): void {
  const direct = normalizeHexColor(value);
  if (direct !== null) {
    set(direct);
    return;
  }
  if (isShorthandHex(value)) {
    const expanded = expandShorthandHex(value);
    set(expanded);
    warnings.push(
      `line ${lineNo}: ${label} shorthand '${value.trim()}' expanded to '${expanded}'`,
    );
    return;
  }
  set(fallback);
  warnings.push(
    `line ${lineNo}: ${label} has malformed hex '${value}'; default '${fallback}' applied`,
  );
}

function tagThemeField(
  config: StyleGuideConfig,
  warnings: string[],
  tag: StyleGuideCalloutTag,
  label: string,
  value: string,
  lineNo: number,
): void {
  const v = value.trim().toLowerCase();
  if (isCalloutTheme(v)) {
    (config.callouts.tagThemes ??= {})[tag] = v;
    return;
  }
  warnings.push(`line ${lineNo}: unknown callout theme '${value}' for ${label}; default theme applied`);
}

function boolField(
  warnings: string[],
  set: (v: boolean) => void,
  fallback: boolean,
  label: string,
  value: string,
  lineNo: number,
): void {
  const parsed = parseBoolean(value);
  if (parsed !== null) {
    set(parsed);
    return;
  }
  set(fallback);
  warnings.push(
    `line ${lineNo}: ${label} has non-boolean '${value}'; default '${String(fallback)}' applied`,
  );
}

function applyKey(
  config: StyleGuideConfig,
  warnings: string[],
  section: string,
  rawKey: string,
  value: string,
  lineNo: number,
): void {
  const key = normKey(rawKey);
  const unknown = (): void => {
    warnings.push(`line ${lineNo}: unrecognized key '${rawKey}' in '${section}' ignored`);
  };

  switch (section) {
    case 'palette': {
      const d = createDefaultStyleGuide().palette;
      switch (key) {
        case 'primary':
          hexField(config, warnings, (v) => (config.palette.primary = v), d.primary, 'palette.primary', value, lineNo);
          return;
        case 'secondary':
          hexField(config, warnings, (v) => (config.palette.secondary = v), d.secondary as string, 'palette.secondary', value, lineNo);
          return;
        case 'accent':
          hexField(config, warnings, (v) => (config.palette.accent = v), d.accent as string, 'palette.accent', value, lineNo);
          return;
        case 'neutraldark':
          hexField(config, warnings, (v) => (config.palette.neutralDark = v), d.neutralDark as string, 'palette.neutralDark', value, lineNo);
          return;
        case 'neutrallight':
          hexField(config, warnings, (v) => (config.palette.neutralLight = v), d.neutralLight as string, 'palette.neutralLight', value, lineNo);
          return;
        case 'chartpalette': {
          const parts = value.split(',').map((p) => p.trim()).filter((p) => p !== '');
          const valid: string[] = [];
          const bad: string[] = [];
          for (const part of parts) {
            const direct = normalizeHexColor(part);
            if (direct !== null) {
              valid.push(direct);
              continue;
            }
            if (isShorthandHex(part)) {
              const expanded = expandShorthandHex(part);
              valid.push(expanded);
              warnings.push(
                `line ${lineNo}: chart palette shorthand '${part}' expanded to '${expanded}'`,
              );
              continue;
            }
            bad.push(part);
          }
          if (bad.length > 0) {
            warnings.push(
              `line ${lineNo}: chart palette entries malformed and dropped (${bad.map((b) => `'${b}'`).join(', ')})`,
            );
          }
          if (valid.length === 0) {
            config.palette.chartPalette = [...d.chartPalette];
            warnings.push(`line ${lineNo}: chart palette has no valid colors; default palette applied`);
          } else if (valid.length < 4) {
            const padded = [...valid];
            const present = new Set(valid.map((c) => c.toUpperCase()));
            for (const c of d.chartPalette) {
              if (padded.length >= 4) break;
              if (present.has(c.toUpperCase())) continue;
              padded.push(c);
              present.add(c.toUpperCase());
            }
            if (padded.length < 4) {
              warnings.push(
                `line ${lineNo}: chart palette fallback colors exhausted; duplicating defaults to reach 4 entries`,
              );
              let i = 0;
              while (padded.length < 4) {
                padded.push(d.chartPalette[i % d.chartPalette.length]);
                i += 1;
              }
            }
            config.palette.chartPalette = padded;
            warnings.push(
              `line ${lineNo}: chart palette has ${valid.length} valid colors (< 4); padded with defaults`,
            );
          } else {
            config.palette.chartPalette = valid;
          }
          return;
        }
        default:
          unknown();
          return;
      }
    }
    case 'tables': {
      const d = createDefaultStyleGuide().tables;
      switch (key) {
        case 'headerfill':
          hexField(config, warnings, (v) => (config.tables.headerFill = v), d.headerFill as string, 'tables.headerFill', value, lineNo);
          return;
        case 'headertext':
          hexField(config, warnings, (v) => (config.tables.headerText = v), d.headerText as string, 'tables.headerText', value, lineNo);
          return;
        case 'headerbold':
          boolField(warnings, (v) => (config.tables.headerBold = v), d.headerBold as boolean, 'tables.headerBold', value, lineNo);
          return;
        case 'alternatingfill':
          hexField(config, warnings, (v) => (config.tables.alternatingFill = v), d.alternatingFill as string, 'tables.alternatingFill', value, lineNo);
          return;
        case 'bordercolor':
          hexField(config, warnings, (v) => (config.tables.borderColor = v), d.borderColor as string, 'tables.borderColor', value, lineNo);
          return;
        case 'keeptogether':
          boolField(warnings, (v) => (config.tables.keepTogether = v), d.keepTogether as boolean, 'tables.keepTogether', value, lineNo);
          return;
        default:
          unknown();
          return;
      }
    }
    case 'typography': {
      const d = createDefaultStyleGuide().typography;
      switch (key) {
        case 'fontfamily':
          if (value !== '') config.typography.fontFamily = value;
          else {
            config.typography.fontFamily = d.fontFamily;
            warnings.push(`line ${lineNo}: typography.fontFamily is empty; default kept`);
          }
          return;
        case 'codefont':
          if (value !== '') config.typography.codeFont = value;
          else {
            config.typography.codeFont = d.codeFont;
            warnings.push(`line ${lineNo}: typography.codeFont is empty; default kept`);
          }
          return;
        case 'heading1':
        case 'heading2': {
          const slot = key as 'heading1' | 'heading2';
          const label = `typography.${slot}`;
          const parts = value.split(',').map((p) => p.trim()).filter((p) => p !== '');
          let color: string | undefined;
          let bold = false;
          for (const part of parts) {
            if (/^bold$/i.test(part)) {
              bold = true;
              continue;
            }
            if (/^(not\s+bold|normal|regular)$/i.test(part)) {
              bold = false;
              continue;
            }
            if (color !== undefined) {
              warnings.push(`line ${lineNo}: ${label} has extra fragment '${part}' ignored`);
              continue;
            }
            const direct = normalizeHexColor(part);
            if (direct !== null) {
              color = direct;
              continue;
            }
            if (isShorthandHex(part)) {
              color = expandShorthandHex(part);
              warnings.push(
                `line ${lineNo}: ${label} shorthand '${part}' expanded to '${color}'`,
              );
              continue;
            }
            warnings.push(
              `line ${lineNo}: ${label} has malformed hex '${part}'; fragment ignored`,
            );
          }
          if (color === undefined) {
            warnings.push(`line ${lineNo}: ${label} has no parseable color ('${value}'); heading style unset`);
            return;
          }
          config.typography[slot] = { color, bold };
          return;
        }
        default:
          unknown();
          return;
      }
    }
    case 'callouts': {
      const d = createDefaultStyleGuide().callouts;
      switch (key) {
        case 'defaulttheme': {
          const v = value.trim().toLowerCase();
          if (isCalloutTheme(v)) config.callouts.defaultTheme = v;
          else {
            config.callouts.defaultTheme = d.defaultTheme;
            warnings.push(`line ${lineNo}: unknown callout theme '${value}'; default '${d.defaultTheme}' applied`);
          }
          return;
        }
        case 'preferredthemes': {
          const parts = value.split(',').map((p) => p.trim().toLowerCase()).filter((p) => p !== '');
          const good = parts.filter(isCalloutTheme);
          const bad = parts.filter((p) => !isCalloutTheme(p));
          for (const b of bad) {
            warnings.push(`line ${lineNo}: unknown callout theme '${b}' dropped from preferredThemes`);
          }
          config.callouts.preferredThemes = good.length > 0 ? good : [...(d.preferredThemes as StyleGuideCalloutTheme[])];
          if (good.length === 0) {
            warnings.push(`line ${lineNo}: preferredThemes has no known themes; defaults applied`);
          }
          return;
        }
        case 'notetheme':
          tagThemeField(config, warnings, 'note', 'Note Theme', value, lineNo);
          return;
        case 'tiptheme':
          tagThemeField(config, warnings, 'tip', 'Tip Theme', value, lineNo);
          return;
        case 'importanttheme':
          tagThemeField(config, warnings, 'important', 'Important Theme', value, lineNo);
          return;
        case 'warningtheme':
          tagThemeField(config, warnings, 'warning', 'Warning Theme', value, lineNo);
          return;
        case 'cautiontheme':
          tagThemeField(config, warnings, 'caution', 'Caution Theme', value, lineNo);
          return;
        default:
          unknown();
          return;
      }
    }
    case 'codeBlocks': {
      const d = createDefaultStyleGuide().codeBlocks;
      switch (key) {
        case 'background':
          hexField(config, warnings, (v) => (config.codeBlocks.background = v), d.background as string, 'codeBlocks.background', value, lineNo);
          return;
        case 'bordercolor':
          hexField(config, warnings, (v) => (config.codeBlocks.borderColor = v), d.borderColor as string, 'codeBlocks.borderColor', value, lineNo);
          return;
        default:
          unknown();
          return;
      }
    }
    case 'layout': {
      const d = createDefaultStyleGuide().layout;
      switch (key) {
        case 'margin': {
          const v = value.trim().toLowerCase();
          if (isMargin(v)) config.layout.margin = v;
          else {
            config.layout.margin = d.margin;
            warnings.push(`line ${lineNo}: unknown margin '${value}'; default '${d.margin}' applied`);
          }
          return;
        }
        case 'preventorphanheadings':
          boolField(warnings, (v) => (config.layout.preventOrphanHeadings = v), d.preventOrphanHeadings as boolean, 'layout.preventOrphanHeadings', value, lineNo);
          return;
        case 'widowcontrol':
          boolField(warnings, (v) => (config.layout.widowControl = v), d.widowControl as boolean, 'layout.widowControl', value, lineNo);
          return;
        default:
          unknown();
          return;
      }
    }
    default:
      warnings.push(`line ${lineNo}: key '${rawKey}' in unknown section ignored`);
  }
}
