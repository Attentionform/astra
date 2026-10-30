/**
 * Visual review prompt law.
 *
 * Pure, dependency-free constant. Canonical home is
 * `packages/engine/src/visual/prompt.ts`.
 */

/**
 * Prompt law: the rendered image is an approximation (LibreOffice metrics
 * differ from Word), the OOXML DOM is the source of truth. Flag only
 * large, user-visible defects — never micro-typography.
 *
 * Contains the required phrases: "image approximate", "XML truth",
 * "large defects only" (asserted by tests).
 */
export const VISUAL_PROMPT_LAW =
  'Image is approximate, XML is truth — flag only large defects (missing chart, overflow table, blank page, orphan heading), never micro-typography. Large defects only.';

/** Large defect kinds the telemetry + vision pass may flag. */
export const VISUAL_LARGE_DEFECT_KINDS = [
  'missing-chart',
  'overflow-table',
  'blank-page',
  'orphan-heading',
] as const;

export type VisualLargeDefectKind = (typeof VISUAL_LARGE_DEFECT_KINDS)[number];

/**
 * Verification rule: never auto-edit from pixels alone — every visual
 * finding must be re-verified against the OOXML DOM before any mutation.
 */
export const VISUAL_VERIFY_RULE =
  'Never auto-edit from the image alone. Re-verify every visual finding against the OOXML DOM before applying any operation.';
