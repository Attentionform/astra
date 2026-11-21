// Supported entry point: `@attentionform/astra-engine/history`.
// The recorded history of a document: start it, record a state, list the states, read one back.
// Reading an older version as a document is shown in examples/09-history-older-version.mjs.
export {
  getRepoDir,
  initAstraRepo,
  commitChange,
  getHistory,
  getHistoryCommitCount,
  getHistoryInitializationState,
  checkoutCommit,
  trackedPredicateFor,
  assertHistoryReadyForWrite,
  INCOMPLETE_SNAPSHOT_WARNING,
  HISTORY_DIR_ENV,
  HISTORY_MAX_COMMITS_ENV,
  HISTORY_MAX_SIDECAR_BYTES_ENV,
  HISTORY_QUOTA_CODE,
} from '../git';
export type { CommitPackage, HistoryQuotaLimits, HistoryInitializationState } from '../git';
