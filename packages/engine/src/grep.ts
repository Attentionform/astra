/**
 * Grep-based text search for document nodes
 * Fast, simple alternative to semantic search for structured documents
 */

export interface GrepOptions {
  regex?: boolean;
  caseSensitive?: boolean;
}

export interface GrepResult {
  id: string;
  preview: string;
  matchContext: string;
}

/**
 * Search nodes for text matches
 * @param nodes - Array of nodes with id, text, and preview
 * @param query - Search query (text or regex pattern)
 * @param options - Search options (regex, caseSensitive)
 * @returns Array of matching nodes with context
 */
export function grepNodes(
  nodes: { id: string; text: string; preview: string }[],
  query: string,
  options: GrepOptions = {}
): GrepResult[] {
  // Don't use 'g' flag for test() - it causes issues with lastIndex
  const flags = options.caseSensitive ? '' : 'i';
  
  let pattern: RegExp;
  try {
    pattern = options.regex 
      ? new RegExp(query, flags)
      : new RegExp(escapeRegex(query), flags);
  } catch {
    // Invalid regex, return empty results
    return [];
  }
  
  return nodes
    .filter(n => pattern.test(n.text))
    .map(n => ({
      id: n.id,
      preview: n.preview,
      matchContext: extractMatchContext(n.text, pattern, 100)
    }));
}

/**
 * Extract context around the first match
 * @param text - Full text content
 * @param pattern - Regex pattern to match
 * @param contextChars - Number of characters to show before/after match
 * @returns Context string with match highlighted
 */
function extractMatchContext(text: string, pattern: RegExp, contextChars: number): string {
  // Reset regex state
  pattern.lastIndex = 0;
  
  const match = pattern.exec(text);
  if (!match) return '';
  
  const start = Math.max(0, match.index - contextChars);
  const end = Math.min(text.length, match.index + match[0].length + contextChars);
  
  let context = text.slice(start, end);
  
  // Add ellipsis if truncated
  if (start > 0) context = '...' + context;
  if (end < text.length) context = context + '...';
  
  return context;
}

/**
 * Escape special regex characters for literal matching
 * @param str - String to escape
 * @returns Escaped string safe for use in RegExp
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
