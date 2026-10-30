import { Marked } from 'marked';

// Block math: "$$" fence lines, a one-line "$$ x $$", or a "\[" / "\]" fence.
// The token text is the untouched LaTeX; code fences are left alone, and a
// "$$" that does not start a line stays a paragraph. A block body never
// crosses a blank line and is at most MAX_BLOCK_MATH characters.
//
// Cost: marked clips a paragraph at every index a block extension's start()
// returns, and each clip is expensive (160 KB of "$$ a" lines took 7 s even
// when start() returned a constant). So start() reports only line starts that
// really open a block-math token; a long run of "$$" or "\[" lines that never
// close is then one ordinary paragraph, and the unclosed fence is reported by
// the compiler.
export const MAX_BLOCK_MATH = 4000;
const DOLLAR_FENCE = new RegExp(
  `^\\$\\$[ \\t]*\\n((?:(?!\\n[ \\t]*\\n)[\\s\\S]){1,${MAX_BLOCK_MATH}}?)\\n\\$\\$[ \\t]*(?:\\n|$)`
);
const DOLLAR_LINE = /^\$\$[ \t]*([^\n$]+?)[ \t]*\$\$[ \t]*(?:\n|$)/;
const BRACKET_FENCE = new RegExp(
  `^\\\\\\[[ \\t]*\\n?((?:(?!\\\\\\]|\\n[ \\t]*\\n)[\\s\\S]){1,${MAX_BLOCK_MATH}}?)\\n?\\\\\\][ \\t]*(?:\\n|$)`
);
const LINE_START = /^(?:\$\$|\\\[)/gm;

function matchBlockMath(src: string): RegExpExecArray | null {
  return DOLLAR_FENCE.exec(src) ?? DOLLAR_LINE.exec(src) ?? BRACKET_FENCE.exec(src);
}

const blockMath = {
  name: 'blockMath',
  level: 'block' as const,
  start(src: string): number | undefined {
    LINE_START.lastIndex = 0;
    for (let m = LINE_START.exec(src); m !== null; m = LINE_START.exec(src)) {
      if (matchBlockMath(src.slice(m.index, m.index + MAX_BLOCK_MATH + 64)) !== null) return m.index;
    }
    return undefined;
  },
  tokenizer(src: string) {
    const m = matchBlockMath(src);
    if (m) return { type: 'blockMath', raw: m[0], text: m[1].trim() };
    return undefined;
  },
};

const parenMath = {
  name: 'parenMath',
  level: 'inline' as const,
  start(src: string): number | undefined {
    const i = src.indexOf('\\(');
    return i < 0 ? undefined : i;
  },
  tokenizer(src: string) {
    const m = /^\\\(([^\n]{1,500}?)\\\)/.exec(src);
    return m ? { type: 'parenMath', raw: m[0], text: m[1] } : undefined;
  },
};

const dollarMath = {
  name: 'dollarMath',
  level: 'inline' as const,
  start(src: string): number | undefined {
    const i = src.indexOf('$');
    return i < 0 ? undefined : i;
  },
  tokenizer(src: string) {
    const m = /^\$(?!\$)(?!\s)((?:\\\$|[^$\n]){1,500}?)(?<![$\s\\])\$(?!\$)(?!\d)/.exec(src);
    return m ? { type: 'dollarMath', raw: m[0], text: m[1] } : undefined;
  },
};

const instanceDefault = new Marked({ extensions: [blockMath, parenMath] });
const instanceDollar = new Marked({ extensions: [blockMath, parenMath, dollarMath] });

/**
 * Lexes Markdown with the math extensions. A `new Marked({ extensions })`
 * instance ignores its extensions unless the call passes `instance.defaults`,
 * so never call `instance.lexer(src, { gfm: true })` alone.
 */
export function lexMarkdown(src: string, opts: { dollarMath?: boolean } = {}) {
  const inst = opts.dollarMath ? instanceDollar : instanceDefault;
  return inst.lexer(src, { ...inst.defaults, gfm: true, breaks: false });
}
