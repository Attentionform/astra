import { ommlToLatex } from './omml';
import { createWordElement, stampRevision } from './ns';
import { wordDiff } from './track-changes-emitter';

type DiffChunk = { type: 'equal' | 'delete' | 'insert'; text: string };
interface Segment { run: Element; start: number; end: number }
interface Opaque { kind: 'field' | 'math' | 'sdt'; start: number; end: number; elements: Element[] }
export interface RunModel { paragraph: Element; segments: Segment[]; opaque: Opaque[]; text: string }
export interface RunEditOptions { tracked: boolean; author: string; date: string; nextId: () => number; includeEmptyFields?: boolean }

const kids = (el: Element): Element[] => Array.from(el.childNodes).filter((n): n is Element => n.nodeType === 1);
const name = (el: Element): string => el.localName || el.tagName.split(':').pop() || '';
const textOf = (run: Element): string => kids(run).filter(el => name(el) === 't').map(el => el.textContent ?? '').join('');

const SAFE_ATOMS = new Set([
  'tab', 'br', 'cr', 'drawing', 'pict', 'object',
  'footnoteReference', 'endnoteReference', 'commentReference',
  'lastRenderedPageBreak', 'noBreakHyphen', 'softHyphen', 'sym', 'ptab',
  'footnoteRef', 'endnoteRef', 'annotationRef',
  'dayShort', 'monthShort', 'yearShort', 'dayLong', 'monthLong', 'yearLong',
]);

function unsupported(reason: string): never {
  throw Object.assign(new Error(`E_VALIDATION: Cannot edit this paragraph safely: ${reason}. Remedy: edit around the affected structure or replace the paragraph explicitly.`), {code: 'E_VALIDATION'});
}

/** Build a text coordinate map without mutating the document. */
export function buildRunModel(paragraph: Element, expectedText: string): RunModel {
  const segments: Segment[] = [];
  const opaque: Opaque[] = [];
  let text = '';
  let field: Opaque | null = null;
  const safeAtoms = SAFE_ATOMS;
  const wordText = (el: Element): string => Array.from(el.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't')).map(t => t.textContent ?? '').join('');
  function visit(container: Element): void {
    for (const child of kids(container)) {
      const tag = name(child);
      if (tag === 'r') {
        const parts = kids(child).filter(el => name(el) !== 'rPr');
        const begin = parts.some(el => name(el) === 'fldChar' && el.getAttribute('w:fldCharType') === 'begin');
        const end = parts.some(el => name(el) === 'fldChar' && el.getAttribute('w:fldCharType') === 'end');
        if (begin) {
          if (field) unsupported('nested fields are unsupported');
          field = {kind:'field', start:text.length, end:text.length, elements:[]};
        }
        if (field) {
          field.elements.push(child);
          text += parts.filter(el => name(el) === 't').map(el => el.textContent ?? '').join('');
          if (end) { field.end = text.length; opaque.push(field); field = null; }
          continue;
        }
        if (parts.some(el => name(el) !== 't' && !safeAtoms.has(name(el)))) unsupported('a run contains an unsupported field marker');
        const content = parts.filter(el => name(el) === 't').map(el => el.textContent ?? '').join('');
        if (content) segments.push({run:child, start:text.length, end:text.length + content.length});
        text += content;
      } else if (tag === 'fldSimple' || tag === 'sdt') {
        const value = wordText(child);
        opaque.push({kind:tag === 'sdt' ? 'sdt' : 'field', start:text.length, end:text.length + value.length, elements:[child]});
        text += value;
      } else if (child.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/math' && (tag === 'oMath' || tag === 'oMathPara')) {
        const latex = ommlToLatex(child) ?? Array.from(child.getElementsByTagName('m:t')).map(t => t.textContent || '').join('');
        const value = `\\(${latex}\\)`;
        opaque.push({kind:'math', start:text.length, end:text.length + value.length, elements:[child]});
        text += value;
      } else if (tag === 'hyperlink' || tag === 'smartTag' || tag === 'ins') {
        visit(child);
      } else if (tag === 'del' || tag === 'pPr' || tag === 'proofErr' || tag.startsWith('bookmark') || tag.startsWith('commentRange')) {
        continue;
      } else {
        unsupported(`unsupported ${child.tagName} structure`);
      }
    }
  }
  visit(paragraph);
  if (field) unsupported('a field has no end marker');
  if (text !== expectedText) unsupported('the paragraph text and text map differ');
  return {paragraph, segments, opaque, text};
}

/** Trim identical whole tokens before the existing bounded word diff. */
export function diffForEdit(oldText: string, newText: string): DiffChunk[] {
  const oldTokens = oldText.split(/(\s+)/);
  const newTokens = newText.split(/(\s+)/);
  let prefix = 0;
  while (prefix < oldTokens.length && prefix < newTokens.length && oldTokens[prefix] === newTokens[prefix]) prefix++;
  let oldEnd = oldTokens.length;
  let newEnd = newTokens.length;
  while (oldEnd > prefix && newEnd > prefix && oldTokens[oldEnd - 1] === newTokens[newEnd - 1]) { oldEnd--; newEnd--; }
  const oldMiddle = oldTokens.slice(prefix, oldEnd).join('');
  const newMiddle = newTokens.slice(prefix, newEnd).join('');
  if (oldMiddle.split(/(\s+)/).length * newMiddle.split(/(\s+)/).length > 2_500_000) unsupported('the changed span is too large; split the edit');
  const result: DiffChunk[] = [];
  const left = oldTokens.slice(0, prefix).join('');
  const right = oldTokens.slice(oldEnd).join('');
  if (left) result.push({type:'equal', text:left});
  result.push(...wordDiff(oldMiddle, newMiddle));
  if (right) result.push({type:'equal', text:right});
  return result;
}

function setText(run: Element, text: string, deleted = false): void {
  const doc = run.ownerDocument!;
  for (const child of kids(run)) if (name(child) === 't' || name(child) === 'delText') run.removeChild(child);
  const element = createWordElement(doc, deleted ? 'w:delText' : 'w:t');
  if (/^\s|\s$/.test(text)) element.setAttribute('xml:space', 'preserve');
  element.appendChild(doc.createTextNode(text));
  run.appendChild(element);
}

/** Split mixed text/object runs at child boundaries while retaining each atom. */
function separateMixedRuns(model: RunModel, changes: readonly {type: DiffChunk['type']; text: string; at: number}[]): void {
  for (const segment of [...model.segments]) {
    const run = segment.run;
    const parts = kids(run).filter(el => name(el) !== 'rPr');
    if (!parts.some(el => name(el) !== 't')) continue;
    const touched = changes.some(chunk => chunk.type === 'insert'
      ? chunk.at >= segment.start && chunk.at <= segment.end
      : chunk.type === 'delete' && chunk.at < segment.end && chunk.at + chunk.text.length > segment.start);
    if (!touched) continue;
    const parent = run.parentNode;
    if (!parent) continue;
    const properties = kids(run).find(el => name(el) === 'rPr');
    for (const child of parts) {
      const next = run.cloneNode(false) as Element;
      if (properties) next.appendChild(properties.cloneNode(true));
      next.appendChild(child);
      parent.insertBefore(next, run);
    }
    parent.removeChild(run);
  }
  const refreshed = buildRunModel(model.paragraph, model.text);
  model.segments = refreshed.segments;
  model.opaque = refreshed.opaque;
}

function splitAt(segments: Segment[], offset: number): void {
  const i = segments.findIndex(segment => segment.start < offset && offset < segment.end);
  if (i < 0) return;
  const segment = segments[i];
  const current = textOf(segment.run);
  const index = offset - segment.start;
  const clone = segment.run.cloneNode(true) as Element;
  setText(segment.run, current.slice(0, index));
  setText(clone, current.slice(index));
  segment.run.parentNode!.insertBefore(clone, segment.run.nextSibling);
  segments.splice(i + 1, 0, {run:clone, start:offset, end:segment.end});
  segment.end = offset;
}

function stamp(element: Element, options: RunEditOptions, id: number): void {
  stampRevision(element, id, options.author, options.date);
}

/** Apply a prepared edit after verifying its old-text coordinates. */
export function applyChunks(model: RunModel, chunks: DiffChunk[], options: RunEditOptions): number[] {
  const expected = chunks.filter(chunk => chunk.type !== 'insert').map(chunk => chunk.text).join('');
  if (expected !== model.text) unsupported('the requested text spans do not match the paragraph');
  const positions: {type: DiffChunk['type']; text: string; at: number}[] = [];
  let position = 0;
  for (const chunk of chunks) {
    positions.push({...chunk, at:position});
    if (chunk.type !== 'insert') position += chunk.text.length;
  }
  for (const chunk of positions) {
    if (chunk.type !== 'equal') for (let i = 0; i < chunk.text.length; i++) {
      const code = chunk.text.charCodeAt(i);
      if (code >= 0xD800 && code <= 0xDBFF) {
        const next = chunk.text.charCodeAt(++i);
        if (!(next >= 0xDC00 && next <= 0xDFFF)) unsupported('the changed text contains an unpaired Unicode surrogate');
      } else if (code >= 0xDC00 && code <= 0xDFFF) unsupported('the changed text contains an unpaired Unicode surrogate');
    }
    const end = chunk.at + (chunk.type === 'insert' ? 0 : chunk.text.length);
    for (const at of [chunk.at, end]) {
      if (at > 0 && at < model.text.length && /[\uD800-\uDBFF]/.test(model.text[at - 1]) && /[\uDC00-\uDFFF]/.test(model.text[at])) unsupported('an edit would split a Unicode surrogate pair');
    }
  }
  for (const chunk of positions) {
    if (chunk.type === 'equal') continue;
    const end = chunk.at + (chunk.type === 'delete' ? chunk.text.length : 0);
    for (const item of model.opaque) {
      const cuts = chunk.type === 'insert'
        ? chunk.at > item.start && chunk.at < item.end
        : chunk.at < item.end && end > item.start;
      const emptyField = chunk.type === 'delete' && options.includeEmptyFields && item.kind === 'field' &&
        item.start === item.end && chunk.at <= item.start && end >= item.end;
      if (!cuts && !emptyField) continue;
      const wholeField = chunk.type === 'delete' && item.kind === 'field' && chunk.at <= item.start && end >= item.end;
      if (!wholeField) unsupported(item.kind === 'math' ? 'paragraph contains math in the changed span' : `the changed span cuts through ${item.kind}`);
      if (options.tracked && item.elements.some(el => name(el) === 'fldSimple'))
        unsupported('tracked deletion of a simple field cannot safely preserve its field structure');
    }
  }
  if (!positions.some(chunk => chunk.type !== 'equal' && (chunk.text.length || (chunk.type === 'delete' && options.includeEmptyFields && model.opaque.some(item => item.kind === 'field' && item.start === item.end))))) return [];
  // A paragraph whose visible text is entirely in one plain text run can be
  // edited in place. This keeps its surrounding anchors and existing run
  // structure without fragmenting a replacement into separate word runs.
  if (!options.tracked && model.segments.length === 1 && model.opaque.length === 0) {
    const run = model.segments[0].run;
    const contents = kids(run).filter(el => name(el) !== 'rPr');
    const textEls = contents.filter(el => name(el) === 't');
    const nonTextEls = contents.filter(el => name(el) !== 't');
    if (textEls.length === 1 && nonTextEls.every(el => SAFE_ATOMS.has(name(el)))) {
      const value = chunks.filter(chunk => chunk.type !== 'delete').map(chunk => chunk.text).join('');
      textEls[0].textContent = value;
      if (/^\s|\s$/.test(value)) textEls[0].setAttribute('xml:space', 'preserve');
      else textEls[0].removeAttribute('xml:space');
      return [];
    }
  }
  const fieldInDelete = (item: Opaque, start: number, end: number): boolean => item.kind === 'field' &&
    item.start >= start && item.end <= end &&
    (item.start !== item.end || options.includeEmptyFields === true || (item.start > start && item.end < end));
  // Reserve ids in document order before changing the tree. Applying spans
  // right to left must not make their public ids run backwards.
  const planned = new Map<typeof positions[number], number[]>();
  const ids: number[] = [];
  if (options.tracked) {
    for (const chunk of positions) {
      if (chunk.type === 'equal' || (!chunk.text.length && !(chunk.type === 'delete' && options.includeEmptyFields))) continue;
      const end = chunk.at + (chunk.type === 'delete' ? chunk.text.length : 0);
      const textCount = model.segments.reduce((total, segment) => {
        const parts = kids(segment.run).filter(el => name(el) !== 'rPr');
        if (!parts.some(el => name(el) !== 't')) return total + Number(segment.start < end && segment.end > chunk.at);
        let offset = segment.start;
        for (const part of parts) if (name(part) === 't') {
          const next = offset + (part.textContent ?? '').length;
          if (offset < end && next > chunk.at) total++;
          offset = next;
        }
        return total;
      }, 0);
      const count = chunk.type === 'insert' ? 1 :
        textCount +
        model.opaque.filter(o => fieldInDelete(o, chunk.at, end))
          .reduce((sum, o) => sum + o.elements.length, 0);
      const reserved = Array.from({length:count}, () => options.nextId());
      planned.set(chunk, reserved);
      ids.push(...reserved);
    }
  }
  separateMixedRuns(model, positions);
  const segments = model.segments;
  for (const chunk of positions.reverse()) {
    if (chunk.type === 'equal') continue;
    const reserved = planned.get(chunk) ?? [];
    let used = 0;
    if (chunk.type === 'delete') {
      const end = chunk.at + chunk.text.length;
      splitAt(segments, chunk.at);
      splitAt(segments, end);
      const fieldElements = model.opaque
        .filter(value => fieldInDelete(value, chunk.at, end))
        .flatMap(value => value.elements.map(element => ({element, field:true})));
      const textElements = segments
        .filter(s => s.start >= chunk.at && s.end <= end && s.end > s.start)
        .map(s => ({element:s.run, field:false}));
      const order = new Map(Array.from(model.paragraph.getElementsByTagName('*')).map((el, i) => [el, i]));
      const targets = [...fieldElements, ...textElements].sort((a,b) =>
        (order.get(a.element) ?? -1) - (order.get(b.element) ?? -1)
      );
      for (const target of targets) {
        const element = target.element;
        const parent = element.parentNode;
        if (!parent) continue;
        if (options.tracked) {
          const wrapper = createWordElement(model.paragraph.ownerDocument!, 'w:del');
          const id = reserved[used++]; stamp(wrapper, options, id);
          parent.insertBefore(wrapper, element);
          wrapper.appendChild(element);
          if (target.field) {
            for (const [from, to] of [['w:t', 'w:delText'], ['w:instrText', 'w:delInstrText']]) {
              for (const source of Array.from(element.getElementsByTagName(from))) {
                const replacement = createWordElement(model.paragraph.ownerDocument!, to);
                for (const attr of Array.from(source.attributes)) replacement.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
                while (source.firstChild) replacement.appendChild(source.firstChild);
                source.parentNode!.replaceChild(replacement, source);
              }
            }
          } else setText(element, textOf(element), true);
        } else parent.removeChild(element);
      }
    } else {
      splitAt(segments, chunk.at);
      const previous = segments.find(s => s.end === chunk.at && s.end > s.start);
      const ref = previous ?? segments.find(s => s.start === chunk.at);
      const doc = model.paragraph.ownerDocument!;
      const run = createWordElement(doc, 'w:r');
      const properties = ref && kids(ref.run).find(el => name(el) === 'rPr');
      if (properties) run.appendChild(properties.cloneNode(true));
      setText(run, chunk.text);
      let inserted: Element = run;
      if (options.tracked) {
        const wrapper = createWordElement(doc, 'w:ins');
        const id = reserved[used++]; stamp(wrapper, options, id);
        wrapper.appendChild(run); inserted = wrapper;
      }
      const fieldAtEnd = model.opaque.find(item => item.end === chunk.at && item.elements[item.elements.length - 1]?.parentNode);
      const fieldAtStart = model.opaque.find(item => item.start === chunk.at && item.elements[0]?.parentNode);
      const anchor = ref?.run.parentNode && name(ref.run.parentNode as Element) === 'del' ? ref.run.parentNode : ref?.run;
      if (fieldAtEnd) {
        const last = fieldAtEnd.elements[fieldAtEnd.elements.length - 1];
        last.parentNode!.insertBefore(inserted, last.nextSibling);
      } else if (fieldAtStart) {
        const first = fieldAtStart.elements[0];
        first.parentNode!.insertBefore(inserted, first);
      } else if (previous && anchor?.parentNode) anchor.parentNode.insertBefore(inserted, anchor.nextSibling);
      else if (anchor?.parentNode) anchor.parentNode.insertBefore(inserted, anchor);
      else model.paragraph.appendChild(inserted);
    }
    if (options.tracked && used !== reserved.length) throw new Error('Internal revision allocation mismatch');
  }
  return ids;
}
