/** The limits that a page must keep. Kept apart from `spec.ts` so that `papers.ts` can use them.
 */
/** The smallest page side that is accepted: 1 in. */
export const MIN_PAGE_SIDE = 1440;
/**
 * The largest page side that is accepted: 22 in (31680 twips). Checked against Word 16 through COM: its page setup refuses more, and a file
 * that holds more opens without an error but is read with the size wrapped at 16 bits. The Open XML validator also rejects it (`w:w` and
 * `w:h` have a maximum of 31680), so ASTra never writes such a page. A file that already has one is left as it is by an edit that does not set a size.
 */
export const MAX_PAGE_SIDE = 31680;
/**
 * The smallest text width that a request may leave after margins and gutter: 0.75 in. A policy of ASTra, informed by Word 16.0.14334, whose
 * object model accepted a text width of 54 pt (0.75 in) and refused 52 pt. Below this Word opens a file but breaks words in the middle.
 */
export const MIN_TEXT_WIDTH = 1080;
/** The smallest text height that a request may leave after margins and gutter: 0.5 in. A policy of ASTra; Word has no limit on text height. */
export const MIN_TEXT_HEIGHT = 720;
/** A margin below this is valid but may be clipped by a printer: 0.25 in. */
export const LOW_MARGIN_WARNING = 360;

