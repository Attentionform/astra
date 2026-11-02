/**
 * Repair rules for known validator error kinds, run by the explicit normalize commands only (never at
 * open or save). Each rule handles one error kind with one unambiguous fix, is pure DOM work, and is
 * idempotent: a second run changes nothing and logs nothing. A rule never changes visible text.
 *
 * The runner (`repairDocxBuffer`) applies the rules one at a time and keeps a rule only when
 *  - the visible text of every part is unchanged,
 *  - the package has no validator error that it did not have before the rule, and
 *  - the rule changed something.
 * A rule that fails a gate is dropped alone and reported in `rolledBack`.
 */
import { createWordElement } from './ns';
import { XMLSerializer } from '@xmldom/xmldom';
import { DOMParser } from './dom-parser';
import AdmZip from '../adm-zip-safe';
import { parsePartStrict } from './normalize';
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, SETTINGS_CHILD_ORDER, TBLPR_CHILD_ORDER, TCPR_CHILD_ORDER, TRPR_CHILD_ORDER } from './schema-order';
import { BASELINE_MAX_ERRORS, buildBaseline, judgeAgainstBaseline, type ValidationBaseline } from './validation-baseline';
import { validateDocxBuffer } from './validator-gate';
import type { ValidationResult } from './validator-backend';

export interface RepairContext {
  /** True when any part of the package marks right-to-left text, which makes start/end alignment ambiguous. */
  hasBidi: boolean;
}

export interface RepairRule {
  id: string;
  description: string;
  /** Does the rule look at this part? */
  appliesTo(part: string): boolean;
  /** Changes the DOM in place and returns one log line per change, each naming its xPath. */
  apply(doc: Document, part: string, ctx: RepairContext): string[];
}

// ── helpers ──────────────────────────────────────────────────────────────

const isElement = (n: Node | null): n is Element => !!n && n.nodeType === 1;

function elementChildren(el: Element): Element[] {
  const out: Element[] = [];
  for (let c = el.firstChild; c; c = c.nextSibling) if (isElement(c)) out.push(c);
  return out;
}

function snapshot(doc: Document | Element, tag: string): Element[] {
  const list = doc.getElementsByTagName(tag);
  const out: Element[] = [];
  for (let i = 0; i < list.length; i++) out.push(list[i] as unknown as Element);
  return out;
}

function xpathOf(el: Element): string {
  const steps: string[] = [];
  let cur: Node | null = el;
  while (isElement(cur)) {
    let n = 1;
    for (let s = cur.previousSibling; s; s = s.previousSibling) if (isElement(s) && s.tagName === cur.tagName) n++;
    steps.unshift(`${cur.tagName}[${n}]`);
    cur = cur.parentNode;
  }
  return '/' + steps.join('/');
}

/**
 * Stable sort of the children whose tag has a rank, among the places those children already occupy.
 * Children without a rank (and text between elements) stay where they are. Returns true when it moved
 * something. Equal ranks keep their order, so valid content is never touched.
 */
function sortByRank(parent: Element, rankOf: (tag: string) => number | undefined): boolean {
  const kids = elementChildren(parent);
  const slots: number[] = [];
  const ranked: Element[] = [];
  kids.forEach((k, i) => {
    if (rankOf(k.tagName) !== undefined) {
      slots.push(i);
      ranked.push(k);
    }
  });
  const sorted = ranked
    .map((el, i) => ({ el, i, r: rankOf(el.tagName) as number }))
    .sort((a, b) => a.r - b.r || a.i - b.i)
    .map((x) => x.el);
  if (sorted.every((el, i) => el === ranked[i])) return false;
  // Write the sorted elements back into the slots: replace each slot's element by a marker, then fill.
  const markers = ranked.map((el) => {
    const m = parent.ownerDocument.createComment('');
    parent.replaceChild(m, el);
    return m;
  });
  markers.forEach((m, i) => parent.replaceChild(sorted[i], m));
  return true;
}

const rankFromOrder = (order: readonly string[]) => {
  const map = new Map(order.map((t, i) => [t, i] as const));
  return (tag: string) => map.get(tag);
};

function tagList(el: Element): string {
  return elementChildren(el).map((c) => c.tagName).join(',');
}

/** Sorts the children of every element named `tag` and logs each element that moved. */
function sortAll(doc: Document, tag: string, rank: (t: string) => number | undefined, label: string): string[] {
  const log: string[] = [];
  for (const el of snapshot(doc, tag)) {
    const before = tagList(el);
    if (sortByRank(el, rank)) log.push(`${label} ${xpathOf(el)}: [${before}] -> [${tagList(el)}]`);
  }
  return log;
}

const isMainOrHdrFtr = (p: string) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(p);

// ── rules ────────────────────────────────────────────────────────────────

const STYLE_CHILD_ORDER = [
  'w:name', 'w:aliases', 'w:basedOn', 'w:next', 'w:link', 'w:autoRedefine', 'w:hidden', 'w:uiPriority', 'w:semiHidden',
  'w:unhideWhenUsed', 'w:qFormat', 'w:locked', 'w:personal', 'w:personalCompose', 'w:personalReply', 'w:rsid', 'w:pPr',
  'w:rPr', 'w:tblPr', 'w:trPr', 'w:tcPr', 'w:tblStylePr',
];

const SECTPR_RANK: Record<string, number> = {
  'w:headerReference': 0, 'w:footerReference': 0, 'w:footnotePr': 1, 'w:endnotePr': 2, 'w:type': 3, 'w:pgSz': 4, 'w:pgMar': 5,
  'w:paperSrc': 6, 'w:pgBorders': 7, 'w:lnNumType': 8, 'w:pgNumType': 9, 'w:cols': 10, 'w:formProt': 11, 'w:vAlign': 12,
  'w:noEndnote': 13, 'w:titlePg': 14, 'w:textDirection': 15, 'w:bidi': 16, 'w:rtlGutter': 17, 'w:docGrid': 18, 'w:printerSettings': 19,
  'w:sectPrChange': 20,
};

const INTEGER_ATTRIBUTES: Record<string, readonly string[]> = {
  'w:pgSz': ['w:w', 'w:h'],
  'w:pgMar': ['w:top', 'w:right', 'w:bottom', 'w:left', 'w:header', 'w:footer', 'w:gutter'],
  'w:tblInd': ['w:w'], 'w:tblW': ['w:w'], 'w:tcW': ['w:w'], 'w:gridCol': ['w:w'], 'w:tblCellSpacing': ['w:w'],
  'w:tab': ['w:pos'], 'w:trHeight': ['w:val'],
  'w:ind': ['w:left', 'w:right', 'w:start', 'w:end', 'w:hanging', 'w:firstLine'],
  'w:spacing': ['w:before', 'w:after', 'w:line'],
};
const MARGIN_SIDES = new Set(['w:top', 'w:left', 'w:bottom', 'w:right', 'w:start', 'w:end']);

const TCMAR_CHILD_RANK: Record<string, number> = {
  'w:top': 0,
  'w:left': 1,
  'w:start': 1,
  'w:bottom': 2,
  'w:right': 3,
  'w:end': 3,
};

const ON_OFF_ONLY = new Set(['w:bidiVisual', 'w:cantSplit', 'w:noWrap', 'w:tblHeader', 'w:qFormat', 'w:unhideWhenUsed', 'w:semiHidden']);

const textOf = (el: Element) => el.textContent ?? '';

export const REPAIR_RULES: readonly RepairRule[] = [
  {
    id: 'settings-order',
    description: 'put the children of w:settings in schema order',
    appliesTo: (p) => p === 'word/settings.xml',
    apply(doc) {
      const root = doc.documentElement;
      if (!root || root.tagName !== 'w:settings') return [];
      const before = tagList(root);
      return sortByRank(root, rankFromOrder(SETTINGS_CHILD_ORDER)) ? [`Reordered w:settings children [${before}] -> [${tagList(root)}]`] : [];
    },
  },
  {
    id: 'style-order',
    description: 'put the children of w:style and their property containers in schema order',
    appliesTo: (p) => p === 'word/styles.xml',
    apply(doc) {
      const log = sortAll(doc, 'w:style', rankFromOrder(STYLE_CHILD_ORDER), 'Reordered');
      for (const [tag, order] of [['w:pPr', PPR_CHILD_ORDER], ['w:rPr', RPR_CHILD_ORDER], ['w:tblPr', TBLPR_CHILD_ORDER], ['w:trPr', TRPR_CHILD_ORDER], ['w:tcPr', TCPR_CHILD_ORDER]] as const) {
        log.push(...sortAll(doc, tag, rankFromOrder(order), 'Reordered'));
      }
      return log;
    },
  },
  {
    id: 'sectpr-order',
    description: 'put the children of w:sectPr in schema order',
    appliesTo: isMainOrHdrFtr,
    apply: (doc) => sortAll(doc, 'w:sectPr', (t) => SECTPR_RANK[t], 'Reordered'),
  },
  {
    id: 'table-children-order',
    description: 'put w:tblPr before w:tblGrid in a table',
    appliesTo: isMainOrHdrFtr,
    apply: (doc) => sortAll(doc, 'w:tbl', (t) => (t === 'w:tblPr' ? 0 : t === 'w:tblGrid' ? 1 : undefined), 'Reordered'),
  },
  {
    id: 'run-properties-first',
    description: 'move w:rPr to the start of its run',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      for (const r of snapshot(doc, 'w:r')) {
        const kids = elementChildren(r);
        const rpr = kids.find((k) => k.tagName === 'w:rPr');
        if (rpr && kids[0] !== rpr) {
          r.insertBefore(rpr, kids[0]);
          log.push(`Moved w:rPr to the start of ${xpathOf(r)}`);
        }
      }
      return log;
    },
  },
  {
    id: 'paragraph-properties-first',
    description: 'move w:pPr to the start of its paragraph',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      for (const p of snapshot(doc, 'w:p')) {
        const kids = elementChildren(p);
        const pprs = kids.filter((k) => k.tagName === 'w:pPr');
        if (pprs.length === 1 && kids[0] !== pprs[0]) {
          p.insertBefore(pprs[0], kids[0]);
          log.push(`Moved w:pPr to the start of ${xpathOf(p)}`);
        }
      }
      return log;
    },
  },
  {
    id: 'integer-attribute',
    description: 'round a decimal in an integer measurement attribute',
    appliesTo: (p) => /^word\/(document|styles|header\d*|footer\d*|footnotes|endnotes|numbering)\.xml$/.test(p),
    apply(doc) {
      const log: string[] = [];
      const fix = (el: Element, attrs: readonly string[]) => {
        for (const a of attrs) {
          const v = el.getAttribute(a);
          if (v !== null && /^-?\d+\.\d+$/.test(v)) {
            const rounded = String(Math.round(Number(v)));
            el.setAttribute(a, rounded);
            log.push(`Rounded ${a}="${v}" to "${rounded}" at ${xpathOf(el)}`);
          }
        }
      };
      const all = doc.getElementsByTagName('*');
      const els: Element[] = [];
      for (let i = 0; i < all.length; i++) els.push(all[i] as unknown as Element);
      for (const el of els) {
        const attrs = INTEGER_ATTRIBUTES[el.tagName];
        if (attrs) fix(el, attrs);
        else if (MARGIN_SIDES.has(el.tagName) && isElement(el.parentNode) && /^w:(tblCellMar|tcMar)$/.test(el.parentNode.tagName)) fix(el, ['w:w']);
      }
      return log;
    },
  },
  {
    id: 'on-off-attribute',
    description: 'write 1/true as on and 0/false as off where the schema accepts only on and off',
    appliesTo: (p) => /^word\/(document|styles|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(p),
    apply(doc) {
      const log: string[] = [];
      for (const tag of ON_OFF_ONLY) {
        for (const el of snapshot(doc, tag)) {
          const v = el.getAttribute('w:val');
          const to = v === '1' || v === 'true' ? 'on' : v === '0' || v === 'false' ? 'off' : null;
          if (to) {
            el.setAttribute('w:val', to);
            log.push(`Rewrote w:val="${v}" as "${to}" at ${xpathOf(el)}`);
          }
        }
      }
      return log;
    },
  },
  {
    id: 'grid-column-type',
    description: 'remove the undeclared w:type attribute of w:gridCol',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      for (const el of snapshot(doc, 'w:gridCol')) {
        if (el.hasAttribute('w:type')) {
          el.removeAttribute('w:type');
          log.push(`Removed w:type from ${xpathOf(el)}`);
        }
      }
      return log;
    },
  },
  {
    id: 'font-family-case',
    description: 'write the font family value Auto as auto',
    appliesTo: (p) => p === 'word/fontTable.xml',
    apply(doc) {
      const log: string[] = [];
      for (const el of snapshot(doc, 'w:family')) {
        if (el.getAttribute('w:val') === 'Auto') {
          el.setAttribute('w:val', 'auto');
          log.push(`Rewrote w:family w:val="Auto" as "auto" at ${xpathOf(el)}`);
        }
      }
      return log;
    },
  },
  {
    id: 'alignment-start-end',
    description: 'write jc start/end as left/right where the text direction is left to right',
    appliesTo: (p) => /^word\/(document|styles|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(p),
    apply(doc, _part, ctx) {
      if (ctx.hasBidi) return [];
      const log: string[] = [];
      for (const el of snapshot(doc, 'w:jc')) {
        const v = el.getAttribute('w:val');
        if (v !== 'start' && v !== 'end') continue;
        const parent = el.parentNode;
        if (!isElement(parent) || !/^w:(pPr|tblPr)$/.test(parent.tagName)) continue;
        const to = v === 'start' ? 'left' : 'right';
        el.setAttribute('w:val', to);
        log.push(`Rewrote w:jc w:val="${v}" as "${to}" at ${xpathOf(el)}`);
      }
      return log;
    },
  },
  {
    id: 'smarttag-unwrap',
    description: 'replace a w:smartTag by its content',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      for (const tag of snapshot(doc, 'w:smartTag')) {
        const parent = tag.parentNode;
        if (!parent) continue;
        for (const kid of Array.from({ length: tag.childNodes.length }, (_, i) => tag.childNodes[i])) {
          if (isElement(kid) && kid.tagName === 'w:smartTagPr') continue;
          parent.insertBefore(kid, tag);
        }
        log.push(`Unwrapped w:smartTag at ${xpathOf(tag)}`);
        parent.removeChild(tag);
      }
      return log;
    },
  },
  {
    id: 'shapetype-duplicate',
    description: 'remove a v:shapetype that repeats an earlier one with the same id and content',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      const ser = new XMLSerializer();
      const seen = new Map<string, string>();
      for (const el of snapshot(doc, 'v:shapetype')) {
        const id = el.getAttribute('id');
        if (!id) continue;
        const xml = ser.serializeToString(el);
        const earlier = seen.get(id);
        if (earlier === undefined) seen.set(id, xml);
        else if (earlier === xml && el.parentNode) {
          log.push(`Removed repeated v:shapetype id="${id}" at ${xpathOf(el)}`);
          el.parentNode.removeChild(el);
        }
      }
      return log;
    },
  },
  {
    id: 'settings-legacy-smarttags',
    description: 'remove w:smartTagType and w:doNotEmbedSmartTags from w:settings, which the validator rejects at every position',
    appliesTo: (p) => p === 'word/settings.xml',
    apply(doc) {
      const root = doc.documentElement;
      if (!root || root.tagName !== 'w:settings') return [];
      const log: string[] = [];
      for (const tag of ['w:smartTagType', 'w:doNotEmbedSmartTags']) {
        for (const el of snapshot(doc, tag)) {
          if (el.parentNode !== root) continue;
          log.push(`Removed ${tag} at ${xpathOf(el)}`);
          root.removeChild(el);
        }
      }
      return log;
    },
  },
  {
    id: 'style-table-defaults',
    description: 'remove table width, layout and bidiVisual from a style when they hold the default value; the validator rejects them in a style',
    appliesTo: (p) => p === 'word/styles.xml',
    apply(doc) {
      const log: string[] = [];
      const isDefault = (el: Element): boolean => {
        switch (el.tagName) {
          case 'w:tblW':
            return (el.getAttribute('w:w') ?? '0') === '0' && ['auto', 'nil', ''].includes(el.getAttribute('w:type') ?? '');
          case 'w:tblLayout':
            return (el.getAttribute('w:type') ?? 'autofit') === 'autofit';
          case 'w:bidiVisual':
            return ['0', 'false', 'off'].includes(el.getAttribute('w:val') ?? '');
          default:
            return false;
        }
      };
      for (const pr of snapshot(doc, 'w:tblPr')) {
        const parent = pr.parentNode;
        if (!parent || !isElement(parent) || parent.tagName !== 'w:style') continue;
        for (let i = pr.childNodes.length - 1; i >= 0; i--) {
          const kid = pr.childNodes[i];
          if (isElement(kid) && isDefault(kid)) {
            log.push(`Removed default ${kid.tagName} at ${xpathOf(kid)}`);
            pr.removeChild(kid);
          }
        }
      }
      return log;
    },
  },
  {
    id: 'paragraph-bare-run-content',
    description: 'put a w:br, w:tab, w:drawing, w:pict or w:sym that sits directly in a paragraph into a run',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      const bare = new Set(['w:br', 'w:tab', 'w:drawing', 'w:pict', 'w:sym']);
      for (const p of snapshot(doc, 'w:p')) {
        for (let i = 0; i < p.childNodes.length; i++) {
          const kid = p.childNodes[i];
          if (!isElement(kid) || !bare.has(kid.tagName)) continue;
          const run = createWordElement(doc, 'w:r');
          p.insertBefore(run, kid);
          p.removeChild(kid);
          run.appendChild(kid);
          log.push(`Put ${kid.tagName} into a run at ${xpathOf(run)}`);
        }
      }
      return log;
    },
  },
  {
    id: 'vml-wrap-inline',
    description: 'remove a w10:wrap whose type is "inline", which is not a wrap type; the shape then has no wrap element',
    appliesTo: isMainOrHdrFtr,
    apply(doc) {
      const log: string[] = [];
      for (const el of snapshot(doc, 'w10:wrap')) {
        if (el.getAttribute('type') !== 'inline' || !el.parentNode) continue;
        log.push(`Removed w10:wrap type="inline" at ${xpathOf(el)}`);
        el.parentNode.removeChild(el);
      }
      return log;
    },
  },
  {
    id: 'color-default-value',
    description: 'give a w:color without w:val the value "auto", which is what a missing value means',
    appliesTo: (p) => /^word\/[^/]+\.xml$/.test(p),
    apply(doc) {
      const log: string[] = [];
      for (const el of snapshot(doc, 'w:color')) {
        if (el.hasAttribute('w:val')) continue;
        el.setAttribute('w:val', 'auto');
        log.push(`Set w:val="auto" on w:color at ${xpathOf(el)}`);
      }
      return log;
    },
  },
  {
    id: 'cell-valign-duplicate',
    description: 'keep one w:vAlign in a w:tcPr: the last, which is the one a renderer applies last',
    appliesTo: (p) => /^word\/[^/]+\.xml$/.test(p),
    apply(doc) {
      const log: string[] = [];
      for (const tcPr of snapshot(doc, 'w:tcPr')) {
        const aligns = elementChildren(tcPr).filter((c) => c.tagName === 'w:vAlign');
        if (aligns.length < 2) continue;
        const kept = aligns[aligns.length - 1];
        for (const extra of aligns.slice(0, -1)) tcPr.removeChild(extra);
        log.push(`Removed ${aligns.length - 1} duplicate w:vAlign from ${xpathOf(tcPr)}, kept w:val="${kept.getAttribute('w:val') ?? ''}"`);
      }
      return log;
    },
  },
  {
    id: 'shading-empty-fill',
    description: 'remove an empty w:fill from w:shd: it is not a colour, and a missing fill means auto',
    appliesTo: (p) => /^word\/[^/]+\.xml$/.test(p),
    apply(doc) {
      const log: string[] = [];
      for (const el of snapshot(doc, 'w:shd')) {
        if (!el.hasAttribute('w:fill') || (el.getAttribute('w:fill') ?? '').trim() !== '') continue;
        el.removeAttribute('w:fill');
        log.push(`Removed the empty w:fill from w:shd at ${xpathOf(el)}`);
      }
      return log;
    },
  },
  {
    id: 'cell-margins-order',
    description: 'put the children of w:tcMar in schema order',
    appliesTo: (p) => /^word\/[^/]+\.xml$/.test(p),
    apply: (doc) => sortAll(doc, 'w:tcMar', (t) => TCMAR_CHILD_RANK[t], 'Reordered'),
  },
];

// ── runner ───────────────────────────────────────────────────────────────

export interface RepairOutcome {
  buffer: Buffer;
  /** One line per change that was kept. */
  normalized: string[];
  /** Validator errors left in the result (0 when the package is valid), or -1 when the check could not run. */
  remaining: number;
  rolledBack: { rule: string; reason: string }[];
}

const REPAIRABLE_PART = /^(word\/[^/]+\.xml|word\/(header|footer)\d*\.xml)$/;

/** Visible text of every part, joined, for the text gate. */
function visibleText(parts: Record<string, string>): string {
  const doc = new DOMParser({ errorHandler: { warning() {}, error() {}, fatalError() {} } });
  const out: string[] = [];
  for (const name of Object.keys(parts).sort()) {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(name)) continue;
    const d = doc.parseFromString(parts[name], 'text/xml');
    for (const tag of ['w:t', 'w:delText', 'a:t', 'm:t']) for (const el of snapshot(d, tag)) out.push(textOf(el));
  }
  return out.join('\u0001');
}

/**
 * Applies the rules one at a time to a package. `validate` defaults to the native validator; it is a
 * parameter so tests can run without it. Never throws for a rule failure: the rule is rolled back.
 */
export async function repairDocxBuffer(
  input: Buffer,
  options: { rules?: readonly RepairRule[]; validate?: (buf: Buffer) => Promise<ValidationResult> } = {},
): Promise<RepairOutcome> {
  const rules = options.rules ?? REPAIR_RULES;
  const validate = options.validate ?? ((b: Buffer) => validateDocxBuffer(b, { maxErrors: BASELINE_MAX_ERRORS }));
  const zip = new AdmZip(input);
  const parts: Record<string, string> = {};
  for (const e of zip.getEntries()) if (!e.isDirectory && REPAIRABLE_PART.test(e.entryName)) parts[e.entryName] = e.getData().toString('utf8');
  const ctx: RepairContext = { hasBidi: Object.values(parts).some((x) => /<w:(bidi|bidiVisual|rtl)\b(?![^>]*w:val="(0|false|off)")/.test(x)) };

  const pack = (p: Record<string, string>): Buffer => {
    const z = new AdmZip(input);
    for (const [name, xml] of Object.entries(p)) if (xml !== parts[name]) z.updateFile(name, Buffer.from(xml, 'utf8'));
    return z.toBuffer();
  };

  let baselineResult: ValidationResult | null = null;
  try {
    baselineResult = await validate(input);
  } catch {
    baselineResult = null;
  }
  let baseline: ValidationBaseline | null = baselineResult ? buildBaseline(baselineResult) : null;
  let remaining = baselineResult ? baselineResult.errors.length : -1;

  let current = { ...parts };
  const currentText = visibleText(current);
  const normalized: string[] = [];
  const rolledBack: { rule: string; reason: string }[] = [];
  const ser = new XMLSerializer();

  for (const rule of rules) {
    const next = { ...current };
    const log: string[] = [];
    try {
      for (const name of Object.keys(current).sort()) {
        if (!rule.appliesTo(name)) continue;
        const doc = parsePartStrict(name, current[name]);
        if (!doc) continue;
        const lines = rule.apply(doc, name, ctx);
        if (lines.length > 0) {
          next[name] = ser.serializeToString(doc);
          log.push(...lines.map((l) => `${name}: ${l}`));
        }
      }
    } catch (err) {
      rolledBack.push({ rule: rule.id, reason: `rule failed: ${err instanceof Error ? err.message.slice(0, 160) : String(err)}` });
      continue;
    }
    if (log.length === 0) continue;
    if (visibleText(next) !== currentText) {
      rolledBack.push({ rule: rule.id, reason: 'visible text would change' });
      continue;
    }
    const candidate = pack(next);
    let result: ValidationResult | null = null;
    try {
      result = await validate(candidate);
    } catch {
      result = null;
    }
    if (result) {
      const verdict = judgeAgainstBaseline(result, baseline ?? { counts: {}, total: 0 });
      if (verdict.kind === 'introduced') {
        rolledBack.push({ rule: rule.id, reason: `adds ${verdict.introduced.length} validator error(s)` });
        continue;
      }
      if (verdict.kind === 'incomparable' && baseline !== null) {
        rolledBack.push({ rule: rule.id, reason: `validator result cannot be compared (${verdict.reason})` });
        continue;
      }
      baseline = buildBaseline(result) ?? baseline;
      remaining = result.errors.length;
    } else if (baseline !== null) {
      rolledBack.push({ rule: rule.id, reason: 'validator did not run' });
      continue;
    }
    current = next;
    normalized.push(...log);
  }
  return { buffer: normalized.length > 0 ? pack(current) : input, normalized, remaining, rolledBack };
}
