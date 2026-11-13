/**
 * Hints for an operation that the schema gate refused: the closest action, key or operation name, and the
 * keys that the action accepts. The schema gate says `Invalid discriminator value` for an action it does not
 * know and `Unrecognized key: "spaceAfter"` for a key it does not know, and neither says what to write instead.
 *
 * The hints read the operation schema itself, so they cannot go out of date with it.
 */
import { OperationSchema } from './operation-schema';

/** The parts of a zod issue that the hints read. */
export interface OperationIssue {
  code: string;
  path: ReadonlyArray<PropertyKey>;
  keys?: ReadonlyArray<string>;
  values?: ReadonlyArray<unknown>;
}

export interface OperationHintOptions {
  /** A sentence that says where to see an example of every action; the caller knows its own tool name. */
  pointer?: string;
}

interface ActionShape {
  action: string;
  keys: string[];
  /** Values of the action's `operation` or `op` field when it is an enum (`table` has many). */
  subOperations?: string[];
}

/** Words people write for an action, mapped to the action. Matched after the name is lowercased. */
const ACTION_WORDS: Record<string, string> = {
  remove: 'delete',
  erase: 'delete',
  del: 'delete',
  drop: 'delete',
  add: 'insert',
  append: 'insert',
  new: 'insert',
  edit: 'replace',
  update: 'replace',
  change: 'replace',
  modify: 'replace',
  rewrite: 'replace',
  format: 'style',
  formatting: 'style',
  font: 'style',
  link: 'hyperlink',
  header: 'headerFooter',
  footer: 'headerFooter',
  headers: 'headerFooter',
  footers: 'headerFooter',
  margins: 'pageSetup',
  margin: 'pageSetup',
  orientation: 'pageSetup',
  pagesetup: 'pageSetup',
  border: 'pageBorder',
  cover: 'coverPage',
  properties: 'metadata',
  tracked: 'trackedChange',
  trackedchanges: 'trackedChange',
  review: 'trackedChange',
  toc: 'field',
  pagenumber: 'field',
  math: 'equation',
  formula: 'equation',
  note: 'footnote',
  graph: 'chart',
  picture: 'image',
  photo: 'image',
  img: 'image',
  bullets: 'list',
  numbering: 'list',
  clear: 'clearFormat',
  clearformatting: 'clearFormat',
  resetformat: 'clearFormat',
};

/** Words people write for a key, mapped to the key, for the keys that are easy to misname. */
export const KEY_WORDS: Record<string, string> = {
  spaceafter: 'spacingAfter',
  spacebefore: 'spacingBefore',
  spacingafterpt: 'spacingAfter',
  paragraphspacing: 'spacingAfter',
  linespace: 'lineSpacing',
  lineheight: 'lineSpacing',
  linespacingmultiple: 'lineSpacing',
  indent: 'indentLeft',
  leftindent: 'indentLeft',
  rightindent: 'indentRight',
  firstlineindent: 'indentFirstLine',
  hangingindent: 'indentHanging',
  align: 'alignment',
  textalign: 'alignment',
  justify: 'alignment',
  font: 'fontFamily',
  fontname: 'fontFamily',
  typeface: 'fontFamily',
  size: 'fontSize',
  textsize: 'fontSize',
  fontcolor: 'color',
  textcolor: 'color',
  colour: 'color',
  background: 'fill',
  backgroundcolor: 'fill',
  shading: 'fill',
  bgcolor: 'fill',
  strike: 'strikethrough',
  strikeout: 'strikethrough',
  bolded: 'bold',
  italics: 'italic',
  underlined: 'underline',
  superscript: 'vertAlign',
  subscript: 'vertAlign',
  node: 'target',
  nodeid: 'target',
  id: 'target',
  content: 'text',
  newtext: 'text',
  href: 'url',
  link: 'url',
  src: 'filePath',
  path: 'filePath',
  file: 'filePath',
  alt: 'altText',
  caption: 'captionText',
};

const MAX_KEYS_LISTED = 60;

let cachedShapes: ActionShape[] | undefined;

/** The enum values of an action's `operation` or `op` field, read from its schema object (empty when it has none). */
function subOperationValues(shape: Record<string, unknown>): string[] {
  for (const key of ['operation', 'op']) {
    let field = shape[key] as { def?: Record<string, unknown> } | undefined;
    // Optional and preprocess wrappers hold the enum one or two levels down.
    for (let depth = 0; depth < 4 && field?.def; depth++) {
      const values = (field.def as { entries?: Record<string, unknown>; values?: unknown[] }).entries;
      if (values && typeof values === 'object') return Object.keys(values);
      field = (field.def['innerType'] ?? field.def['in'] ?? field.def['out']) as { def?: Record<string, unknown> } | undefined;
    }
  }
  return [];
}

/** Verbs people put in front of an action name (`setTable`, `addImage`). Matched after the name is lowercased. */
const ACTION_VERBS = ['set', 'add', 'insert', 'update', 'edit', 'create', 'apply', 'remove', 'delete', 'change', 'make'];

/**
 * Candidates for a name that is a verb plus an action (`setTable` gives `table`). When the verb is itself an
 * action or a word for one (`deleteTable`), that action is a candidate as well, because the name is ambiguous.
 */
function verbPrefixCandidates(given: string, actions: readonly string[]): string[] {
  const key = squash(given);
  for (const verb of ACTION_VERBS) {
    if (!key.startsWith(verb) || key.length - verb.length < 3) continue;
    const rest = key.slice(verb.length);
    const target = actions.find((a) => squash(a) === rest);
    if (!target) continue;
    const out = [target];
    const verbAction = actions.find((a) => squash(a) === verb) ?? ACTION_WORDS[verb];
    if (verbAction && actions.includes(verbAction) && verbAction !== target) out.push(verbAction);
    return out;
  }
  return [];
}
function actionShapes(): ActionShape[] {
  if (cachedShapes) return cachedShapes;
  const shapes: ActionShape[] = [];
  const options = (OperationSchema as unknown as { options: Array<{ shape: Record<string, { def?: { values?: unknown[] } }> }> }).options;
  for (const option of options) {
    const literal = option.shape['action']?.def?.values?.[0];
    if (typeof literal !== 'string') continue;
    const shape: ActionShape = { action: literal, keys: Object.keys(option.shape).filter((k) => k !== 'action') };
    const sub = subOperationValues(option.shape as unknown as Record<string, unknown>);
    if (sub.length > 0) shape.subOperations = sub;
    shapes.push(shape);
  }
  cachedShapes = shapes;
  return shapes;
}

function distance(a: string, b: string): number {
  if (a === b) return 0;
  let previous = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const row = [i];
    for (let j = 1; j <= b.length; j++) {
      row.push(Math.min(previous[j] + 1, row[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)));
    }
    previous = row;
  }
  return previous[b.length];
}

const squash = (name: string): string => name.toLowerCase().replace(/[^a-z0-9]/g, '');

/**
 * The candidates closest to `given`, best first (at most two): a match that ignores case and separators,
 * then a known word for it, then a short edit distance.
 */
export function closestNames(given: string, candidates: readonly string[], words: Record<string, string> = {}): string[] {
  const key = squash(given);
  if (!key) return [];
  const exact = candidates.filter((c) => squash(c) === key);
  if (exact.length > 0) return exact.slice(0, 2);
  const word = words[key];
  if (word && candidates.includes(word)) return [word];
  const limit = Math.max(2, Math.ceil(key.length / 4));
  return candidates
    .map((c) => ({ c, d: distance(key, squash(c)) }))
    .filter((x) => x.d <= limit)
    .sort((x, y) => x.d - y.d || x.c.localeCompare(y.c))
    .slice(0, 2)
    .map((x) => x.c);
}

const quoted = (names: readonly string[]): string => names.map((n) => `"${n}"`).join(' or ');

function operationAt(operations: unknown, index: PropertyKey | undefined): Record<string, unknown> | undefined {
  if (!Array.isArray(operations) || typeof index !== 'number') return undefined;
  const op = operations[index];
  return op && typeof op === 'object' && !Array.isArray(op) ? (op as Record<string, unknown>) : undefined;
}

/** Hints for one issue, or none when the issue is not one of the three kinds that the hints cover. */
function hintFor(operations: unknown, issue: OperationIssue): string | undefined {
  const [root, index, field] = issue.path;
  if (root !== 'operations') return undefined;
  const op = operationAt(operations, index);
  if (!op) return undefined;
  const label = `operation ${String(index)}`;
  const shapes = actionShapes();

  // An action the schema does not know.
  if (field === 'action' && issue.path.length === 3 && issue.code === 'invalid_union') {
    const given = op['action'];
    if (typeof given !== 'string') return undefined;
    const actions = shapes.map((s) => s.action);
    let near = closestNames(given, actions, ACTION_WORDS);
    if (near.length === 0) near = verbPrefixCandidates(given, actions);
    const parts = [`${label}: "${given}" is not an action.`];
    if (near.length > 0) {
      parts.push(`Did you mean ${near.map((n) => `"${n}"`).join(' or ')}?`);
      // What the best candidate accepts, so the caller can see the fields without another call. A second
      // candidate is named only; listing its fields as well would bury the answer.
      const candidate = shapes.find((s) => s.action === near[0]);
      if (candidate) {
        parts.push(`Keys of "${candidate.action}": ${candidate.keys.slice(0, MAX_KEYS_LISTED).join(', ')}${candidate.keys.length > MAX_KEYS_LISTED ? ', ...' : ''}.`);
        if (candidate.subOperations) parts.push(`Operations of "${candidate.action}": ${candidate.subOperations.join(', ')}.`);
      }
    }
    parts.push(`Actions: ${actions.join(', ')}.`);
    return parts.join(' ');
  }

  // A key that the action does not accept.
  if (issue.code === 'unrecognized_keys' && issue.path.length === 2 && issue.keys && issue.keys.length > 0) {
    const action = typeof op['action'] === 'string' ? (op['action'] as string) : undefined;
    const shape = shapes.find((s) => s.action === action);
    if (!shape) return undefined;
    const parts: string[] = [];
    for (const key of issue.keys) {
      const near = closestNames(key, shape.keys, KEY_WORDS);
      parts.push(`${label}: "${key}" is not a key of "${shape.action}".${near.length > 0 ? ` Did you mean ${quoted(near)}?` : ''}`);
    }
    const shown = shape.keys.slice(0, MAX_KEYS_LISTED).join(', ');
    parts.push(`Keys of "${shape.action}": ${shown}${shape.keys.length > MAX_KEYS_LISTED ? ', ...' : ''}.`);
    return parts.join(' ');
  }

  // A sub-operation name that the action does not know (`table` with `addRows`).
  if (issue.code === 'invalid_value' && issue.path.length === 3 && (field === 'operation' || field === 'op') && Array.isArray(issue.values)) {
    const given = op[field as string];
    if (typeof given !== 'string') return undefined;
    const known = issue.values.filter((v): v is string => typeof v === 'string');
    const near = closestNames(given, known);
    const action = typeof op['action'] === 'string' ? (op['action'] as string) : 'this action';
    if (near.length === 0) return undefined;
    return `${label}: "${given}" is not an operation of "${action}". Did you mean ${quoted(near)}?`;
  }
  return undefined;
}

/**
 * What to write instead, for the schema issues of a refused batch: one sentence group per issue that the
 * hints cover, or `undefined` when none is covered. `operations` is the array that was checked.
 */
export function operationIssueRemedy(
  operations: unknown,
  issues: ReadonlyArray<OperationIssue>,
  options: OperationHintOptions = {},
): string | undefined {
  const hints: string[] = [];
  for (const issue of issues.slice(0, 3)) {
    const hint = hintFor(operations, issue);
    if (hint && !hints.includes(hint)) hints.push(hint);
  }
  if (hints.length === 0) return undefined;
  return options.pointer ? `${hints.join(' ')} ${options.pointer}` : hints.join(' ');
}
