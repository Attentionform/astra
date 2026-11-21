import type { NewPageOptions } from '@attentionform/astra-engine';

/** The page flags of `create` and `import`: --page-size, --landscape, --portrait, --margins. */
export interface PageFlags {
  size?: string;
  landscape?: boolean;
  portrait?: boolean;
  margins?: string;
}

export const PAGE_FLAG_USAGE = '--page-size <paper|WxH[unit]> [--landscape|--portrait] [--margins <length|top=..,bottom=..,left=..,right=..,header=..,footer=..,gutter=..>]';

/**
 * Reads one page flag at `args[i]`. Returns null when the argument is not a page flag, the index of the last argument it used, or an error message.
 */
export function readPageFlag(cmd: string, args: string[], i: number, flags: PageFlags): { last: number } | { error: string } | null {
  const a = args[i];
  const valued = (name: string, hint: string): { value: string; last: number } | { error: string } | null => {
    if (a === name) {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) return { error: `${cmd} ${name} requires a value (${hint})` };
      return { value: next, last: i + 1 };
    }
    if (a.startsWith(`${name}=`)) return { value: a.slice(name.length + 1), last: i };
    return null;
  };
  if (a === '--landscape') {
    flags.landscape = true;
    return { last: i };
  }
  if (a === '--portrait') {
    flags.portrait = true;
    return { last: i };
  }
  for (const [name, hint, key] of [
    ['--page-size', 'a paper name such as A4 or Letter, or a size such as 148x210mm', 'size'],
    ['--margins', 'a length such as 20mm, or top=20mm,left=25mm', 'margins'],
  ] as const) {
    const got = valued(name, hint);
    if (got === null) continue;
    if ('error' in got) return got;
    flags[key] = got.value;
    return { last: got.last };
  }
  return null;
}

/** The engine's page option for the flags, or an error message. Undefined when no flag was given. */
export function pageOptionsFrom(cmd: string, flags: PageFlags): NewPageOptions | undefined | { error: string } {
  if (flags.landscape && flags.portrait) return { error: `${cmd}: --landscape and --portrait cannot be used together` };
  const page: NewPageOptions = {};
  if (flags.size !== undefined) page.size = flags.size;
  if (flags.landscape) page.orientation = 'landscape';
  if (flags.portrait) page.orientation = 'portrait';
  if (flags.margins !== undefined) {
    if (flags.margins.includes('=')) {
      const parts: Record<string, string> = {};
      for (const piece of flags.margins.split(',')) {
        const [k, v] = piece.split('=');
        if (!k || v === undefined || v === '') return { error: `${cmd}: --margins needs a length or name=length pairs such as top=20mm,left=25mm (got "${flags.margins}")` };
        parts[k.trim()] = v.trim();
      }
      page.margins = parts as NewPageOptions['margins'];
    } else {
      page.margins = flags.margins;
    }
  }
  return Object.keys(page).length > 0 ? page : undefined;
}
