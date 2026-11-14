/**
 * Runs the engines of a plan in order, with the fallback rules shared by every caller (page images and PDF
 * export, renderers and rasterizers):
 *  - an explicit choice (`word`, `libreoffice`, `windows`, `poppler`) has one engine and never falls back;
 *  - `auto` falls through to the next engine on any failure and says so (`fallbackFrom`, `fallbackReason`);
 *  - an infrastructure failure of a Windows-only engine (Word, the Windows PDF renderer) opens that engine's
 *    circuit breaker, so a long-lived process does not pay a full timeout on every call;
 *  - when every engine failed, the LAST error is thrown (same code, same type) with the earlier failures named
 *    in its message, so neither cause is lost.
 */

import { sanitiseDetail } from './ps-runner';
import { CircuitBreaker, type FailureClass } from './renderer-preference';

export interface PlanOutcome<T, E extends string> {
  value: T;
  used: E;
  fallbackFrom?: E;
  fallbackReason?: string;
}

export function failureClassOf(err: unknown): FailureClass | undefined {
  const c = (err as { failureClass?: unknown } | null)?.failureClass;
  return c === 'infrastructure' || c === 'document' || c === 'request' ? c : undefined;
}

/** One short line for `fallbackReason`: the error's own message, without its code prefix and remedy. */
export function shortReason(err: unknown): string {
  const msg = err instanceof Error ? err.message : String(err);
  return sanitiseDetail(msg.replace(/^\[?E_[A-Z_]+\]?:?\s*/, ''), 200);
}

export interface PlanRules {
  /** More than one engine may be tried (`auto` with several engines available). */
  canFallBack: boolean;
  /** The caller asked for `auto`: only then does an infrastructure failure open a circuit breaker. */
  auto: boolean;
}

export async function runPlan<T, E extends string>(
  engines: readonly E[],
  rules: PlanRules,
  run: (engine: E) => Promise<T>,
  breakers: Partial<Record<E, CircuitBreaker>> = {}
): Promise<PlanOutcome<T, E>> {
  const failures: Array<{ engine: E; err: unknown }> = [];
  for (let i = 0; i < engines.length; i++) {
    const engine = engines[i];
    try {
      const value = await run(engine);
      breakers[engine]?.reset();
      const first = failures[0];
      return {
        value,
        used: engine,
        ...(first ? { fallbackFrom: first.engine, fallbackReason: shortReason(first.err) } : {}),
      };
    } catch (err: unknown) {
      failures.push({ engine, err });
      // Only `auto` consults a breaker, so only an `auto` run may open it (an explicit choice ignores it).
      if (rules.auto && failureClassOf(err) === 'infrastructure') breakers[engine]?.trip();
      if (!rules.canFallBack || i === engines.length - 1) break;
    }
  }
  const last = failures[failures.length - 1];
  if (failures.length > 1 && last.err instanceof Error) {
    const earlier = failures
      .slice(0, -1)
      .map((f) => `${f.engine} failed first: ${shortReason(f.err)}`)
      .join('; ');
    last.err.message = `${last.err.message} (${earlier})`;
    (last.err as { earlierFailures?: string[] }).earlierFailures = failures.slice(0, -1).map((f) => `${f.engine}: ${shortReason(f.err)}`);
  }
  throw last.err;
}
