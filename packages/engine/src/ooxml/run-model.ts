import { ommlToLatex } from './omml';
import { WORD_NS, createWordElement, stampRevision } from './ns';
import { wordDiff } from './track-changes-emitter';

type DiffChunk = { type: 'equal' | 'delete' | 'insert'; text: string };
interface Segment { run: Element; start: number; end: number }
interface Opaque { kind: 'field' | 'math' | 'sdt' | 'object' | 'hidden' | 'ruby'; start: number; end: number; elements: Element[] }
export interface RunModel { paragraph: Element; segments: Segment[]; opaque: Opaque[]; text: string }
export interface RunEditOptions { tracked: boolean; author: string; date: string; nextId: () => number; includeEmptyFields?: boolean;
  /** A tracked delete of the whole paragraph may wrap a run that holds a text box or drawing that contains text. */
  allowObjectDelete?: boolean }

const kids = (el: Element): Element[] => Array.from(el.childNodes).filter((n): n is Element => n.nodeType === 1);
const name = (el: Element): string => el.localName || el.tagName.split(':').pop() || '';
const textOf = (run: Element): string => kids(run).filter(el => name(el) === 't').map(el => el.textContent ?? '').join('');

/**
 * Copy of a run's properties for a run that continues the same text (one half of a split). A tracked formatting
 * change (w:rPrChange) is a revision with its own id, so each copy gets a fresh id and the file keeps unique ids.
 * Author, date and the old properties stay: the same formatting change now covers each part.
 */
function copyProperties(properties: Element, nextId: () => number): Element {
  const copy = properties.cloneNode(true) as Element;
  for (const change of Array.from(copy.childNodes)) {
    if (change.nodeType === 1 && name(change as Element) === 'rPrChange') (change as Element).setAttributeNS(WORD_NS, 'w:id', String(nextId()));
  }
  return copy;
}

/** Copy of a run's properties for new text: the new text was not reformatted by a reviewer, so no formatting change is carried over. */
function propertiesForNewText(properties: Element): Element {
  const copy = properties.cloneNode(true) as Element;
  for (const change of Array.from(copy.childNodes)) {
    if (change.nodeType === 1 && name(change as Element) === 'rPrChange') copy.removeChild(change);
  }
  return copy;
}

const SAFE_ATOMS = new Set([
  'tab', 'br', 'cr', 'drawing', 'pict', 'object',
  'footnoteReference', 'endnoteReference', 'commentReference',
  'AlternateContent', 'lastRenderedPageBreak', 'noBreakHyphen', 'softHyphen', 'sym', 'ptab',
  'footnoteRef', 'endnoteRef', 'annotationRef',
  'dayShort', 'monthShort', 'yearShort', 'dayLong', 'monthLong', 'yearLong',
]);

const FIELD_SPAN = 'this paragraph holds the start or end of a field that spans paragraphs, such as a table of contents; edit text outside the field, or update the field in Word';
const PARTIAL_MOVE = 'this paragraph holds part of a tracked move (a partial move); accept or reject the move with trackedChange first, then edit';

const GENERIC_REMEDY = 'Remedy: edit around the affected structure or replace the paragraph explicitly.';
// Replacing the whole paragraph is refused for these two as well, so the generic remedy would send the caller in a circle.
const INLINE_OBJECT_REMEDY = 'Remedy: edit the text outside the content control or equation, delete the whole paragraph with trackChanges false, or make the change in Word.';
const SIMPLE_FIELD_REMEDY = 'Remedy: edit the text outside the field, or make the change in Word.';
// A changed span that covers only part of a field. Replacing the paragraph again would give the same span, so the remedy is to take in all of the field text.
const FIELD_CUT_REMEDY = 'Remedy: change all of the text of the field in the same edit, or edit the text outside the field.';
const RUBY_REMEDY = 'Remedy: edit the text outside the ruby text, delete the whole paragraph with trackChanges false, or make the change in Word.';

/** Refuses an edit before anything changes. `remedy` is the full sentence; `null` means the reason already says what to do. */
function unsupported(reason: string, remedy: string | null = GENERIC_REMEDY): never {
  throw Object.assign(new Error(`E_VALIDATION: Cannot edit this paragraph safely: ${reason}.${remedy ? ` ${remedy}` : ''}`), {code: 'E_VALIDATION'});
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
        if (parts.some(el => name(el) === 'fldChar' || name(el) === 'instrText')) unsupported(FIELD_SPAN, null);
        if (parts.some(el => name(el) === 'ruby')) {
          // Ruby text (furigana): the stored paragraph text counts the annotation and the base text, so the model counts both
          // as one block that no edit may cut. The text around it can be edited.
          if (parts.some(el => name(el) !== 'ruby' && name(el) !== 't') || parts.some(el => name(el) === 't')) unsupported('a run holds ruby text together with other content; edit the other runs', RUBY_REMEDY);
          const hidden = parts.map(wordText).join('');
          if (hidden) {
            opaque.push({kind:'ruby', start:text.length, end:text.length + hidden.length, elements:[child]});
            text += hidden;
          }
          continue;
        }
        const unknown = parts.find(el => name(el) !== 't' && !safeAtoms.has(name(el)));
        if (unknown) unsupported(`a run contains an element that the editor does not model (w:${name(unknown)})`);
        const content = parts.filter(el => name(el) === 't').map(el => el.textContent ?? '').join('');
        // A text box or drawing keeps its text in w:t elements below the run. The stored paragraph text
        // counts them, so the model counts them too, as one block that no edit may cut.
        const hidden = parts.filter(el => name(el) !== 't').map(wordText).join('');
        if (hidden) {
          if (content) unsupported('a run holds both text and a text box or drawing with text; edit the other runs');
          opaque.push({kind:'object', start:text.length, end:text.length + hidden.length, elements:[child]});
          text += hidden;
          continue;
        }
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
      } else if (tag === 'hyperlink' || tag === 'smartTag' || tag === 'ins' || tag === 'bdo' || tag === 'dir' || tag === 'customXml') {
        // Wrappers that hold ordinary runs. A run that an edit splits stays inside its wrapper, because the split
        // inserts the new run into the run's own parent.
        visit(child);
      } else if (tag === 'AlternateContent') {
        const hidden = wordText(child);
        if (hidden) {
          opaque.push({kind:'object', start:text.length, end:text.length + hidden.length, elements:[child]});
          text += hidden;
        }
      } else if (tag === 'del') {
        // Text of an insertion inside a deletion is still counted by the stored paragraph text.
        const hidden = wordText(child);
        if (hidden) {
          opaque.push({kind:'hidden', start:text.length, end:text.length + hidden.length, elements:[child]});
          text += hidden;
        }
      } else if (tag === 'moveFrom' || tag === 'moveTo' || tag.startsWith('moveFromRange') || tag.startsWith('moveToRange')) {
        unsupported(PARTIAL_MOVE, null);
      } else if (tag === 'pPr' || tag === 'proofErr' || tag.startsWith('bookmark') || tag.startsWith('commentRange') || tag === 'permStart' || tag === 'permEnd') {
        continue;
      } else {
        unsupported(`unsupported ${child.tagName} structure`);
      }
    }
  }
  visit(paragraph);
  if (field) unsupported(FIELD_SPAN, null);
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
function separateMixedRuns(model: RunModel, changes: readonly {type: DiffChunk['type']; text: string; at: number}[], nextId: () => number): void {
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
    for (const [index, child] of parts.entries()) {
      const next = run.cloneNode(false) as Element;
      if (properties) next.appendChild(index === 0 ? properties.cloneNode(true) : copyProperties(properties, nextId));
      next.appendChild(child);
      parent.insertBefore(next, run);
    }
    parent.removeChild(run);
  }
  const refreshed = buildRunModel(model.paragraph, model.text);
  model.segments = refreshed.segments;
  model.opaque = refreshed.opaque;
}

function splitAt(segments: Segment[], offset: number, nextId: () => number): void {
  const i = segments.findIndex(segment => segment.start < offset && offset < segment.end);
  if (i < 0) return;
  const segment = segments[i];
  const current = textOf(segment.run);
  const index = offset - segment.start;
  const clone = segment.run.cloneNode(true) as Element;
  const cloneProps = kids(clone).find(el => name(el) === 'rPr');
  if (cloneProps) clone.replaceChild(copyProperties(cloneProps, nextId), cloneProps);
  setText(segment.run, current.slice(0, index));
  setText(clone, current.slice(index));
  segment.run.parentNode!.insertBefore(clone, segment.run.nextSibling);
  segments.splice(i + 1, 0, {run:clone, start:offset, end:segment.end});
  segment.end = offset;
}

function stamp(element: Element, options: RunEditOptions, id: number): void {
  stampRevision(element, id, options.author, options.date);
}

/**
 * A simple field (`w:fldSimple`) can be converted to a complex field (begin, instruction, separate, result, end) when its
 * result is plain runs. Only then can a tracked deletion be written for it: `w:fldSimple` may not sit inside `w:del`, and a
 * `w:del` inside it would delete the cached result only and leave the field behind.
 */
function simpleFieldCanBeConverted(field: Element): boolean {
  if (!field.getAttributeNS(WORD_NS, 'instr') && !field.getAttribute('w:instr')) return false;
  return kids(field).every(child => name(child) === 'r' &&
    kids(child).every(el => name(el) === 'rPr' || name(el) === 't' || SAFE_ATOMS.has(name(el))));
}

/**
 * Replaces a simple field with the same field written as a complex field and returns the new runs in order. The result
 * runs move over unchanged, so the text, its formatting and the text coordinates stay the same. The marker runs copy the
 * formatting of the first result run (without a tracked formatting change, which belongs to that run only). `w:fldLock`
 * and `w:dirty` move to the begin marker, where a complex field keeps them.
 */
function expandSimpleField(field: Element): Element[] {
  const doc = field.ownerDocument!;
  const parent = field.parentNode!;
  const results = kids(field);
  const properties = results.map(run => kids(run).find(el => name(el) === 'rPr')).find(Boolean);
  const markerRun = (fill: (run: Element) => void): Element => {
    const run = createWordElement(doc, 'w:r');
    if (properties) run.appendChild(propertiesForNewText(properties));
    fill(run);
    return run;
  };
  const fieldChar = (type: 'begin' | 'separate' | 'end'): Element => markerRun(run => {
    const mark = createWordElement(doc, 'w:fldChar');
    mark.setAttributeNS(WORD_NS, 'w:fldCharType', type);
    if (type === 'begin') {
      for (const attribute of ['fldLock', 'dirty']) {
        const value = field.getAttributeNS(WORD_NS, attribute) ?? field.getAttribute(`w:${attribute}`);
        if (value) mark.setAttributeNS(WORD_NS, `w:${attribute}`, value);
      }
    }
    run.appendChild(mark);
  });
  const instruction = markerRun(run => {
    const text = createWordElement(doc, 'w:instrText');
    text.setAttribute('xml:space', 'preserve');
    text.appendChild(doc.createTextNode(field.getAttributeNS(WORD_NS, 'instr') || field.getAttribute('w:instr') || ''));
    run.appendChild(text);
  });
  const expanded = [fieldChar('begin'), instruction, fieldChar('separate'), ...results, fieldChar('end')];
  for (const run of expanded) parent.insertBefore(run, field);
  parent.removeChild(field);
  return expanded;
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
  // Simple fields that a tracked deletion covers whole. They are converted after every check has passed, so a refusal
  // leaves the paragraph as it was.
  const toConvert: Opaque[] = [];
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
      if (item.kind === 'object' || item.kind === 'hidden') {
        const covered = chunk.type === 'delete' && chunk.at <= item.start && end >= item.end;
        if (covered && options.allowObjectDelete) continue;
        unsupported(item.kind === 'object'
          ? 'the changed span would cut or remove a text box or drawing that contains text; edit the text outside it'
          : 'the changed span reaches text inside a deleted revision; accept or reject that revision first');
      }
      if (item.kind === 'ruby') unsupported('the changed span would cut or remove ruby text (furigana)', RUBY_REMEDY);
      const wholeField = chunk.type === 'delete' && item.kind === 'field' && chunk.at <= item.start && end >= item.end;
      if (!wholeField) {
        if (item.kind === 'math' || item.kind === 'sdt') unsupported(item.kind === 'math' ? 'paragraph contains math in the changed span' : 'the changed span cuts through sdt', INLINE_OBJECT_REMEDY);
        unsupported(`the changed span cuts through ${item.kind}`, FIELD_CUT_REMEDY);
      }
      const simple = options.tracked ? item.elements.find(el => name(el) === 'fldSimple') : undefined;
      if (simple) {
        if (!simpleFieldCanBeConverted(simple))
          unsupported('a tracked deletion of a simple field needs the field result to be plain runs, and this one holds a link, a tracked change or another wrapper', SIMPLE_FIELD_REMEDY);
        if (!toConvert.includes(item)) toConvert.push(item);
      }
    }
  }
  if (!positions.some(chunk => chunk.type !== 'equal' && (chunk.text.length || (chunk.type === 'delete' && options.includeEmptyFields && model.opaque.some(item => item.kind === 'field' && item.start === item.end))))) return [];
  for (const item of toConvert) item.elements = expandSimpleField(item.elements.find(el => name(el) === 'fldSimple')!);
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
  const fieldInDelete = (item: Opaque, start: number, end: number): boolean => item.kind === 'hidden' ? false :
    item.kind === 'object' ? item.start >= start && item.end <= end :
    item.kind === 'field' && item.start >= start && item.end <= end &&
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
  separateMixedRuns(model, positions, options.nextId);
  const segments = model.segments;
  for (const chunk of positions.reverse()) {
    if (chunk.type === 'equal') continue;
    const reserved = planned.get(chunk) ?? [];
    let used = 0;
    if (chunk.type === 'delete') {
      const end = chunk.at + chunk.text.length;
      splitAt(segments, chunk.at, options.nextId);
      splitAt(segments, end, options.nextId);
      const fieldElements = model.opaque
        .filter(value => fieldInDelete(value, chunk.at, end))
        .flatMap(value => value.elements.map(element => ({element, field:value.kind === 'field'})));
      const textElements = segments
        .filter(s => s.start >= chunk.at && s.end <= end && s.end > s.start)
        .map(s => ({element:s.run, field:false}));
      const isObject = new Set(model.opaque.filter(value => value.kind === 'object').flatMap(value => value.elements));
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
          } else if (!isObject.has(element)) setText(element, textOf(element), true);
        } else parent.removeChild(element);
      }
    } else {
      splitAt(segments, chunk.at, options.nextId);
      const previous = segments.find(s => s.end === chunk.at && s.end > s.start);
      const ref = previous ?? segments.find(s => s.start === chunk.at);
      const doc = model.paragraph.ownerDocument!;
      const run = createWordElement(doc, 'w:r');
      const properties = ref && kids(ref.run).find(el => name(el) === 'rPr');
      if (properties) run.appendChild(propertiesForNewText(properties));
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
