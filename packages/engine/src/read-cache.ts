/**
 * FileReadCache - Tracks which nodes have been read in the current session
 * Prevents the agent from editing nodes it hasn't read (anti-hallucination)
 *
 * Pattern source: qwen-code/packages/core/src/tools/priorReadEnforcement.ts
 * Bounded LRU: insertion-order Map, oldest evicted past the cap so a
 * long session cannot grow without bound. Content identity is sha256 hex.
 */

import { createHash } from 'crypto';

/** Default cap (~thousands) so long sessions stay bounded. */
export const FILE_READ_CACHE_MAX_ENTRIES = 5000;

export interface ReadEntry {
  lastReadAt: number;
  wasCacheable: boolean;
  contentHash: string;
}

export class FileReadCache {
  private readNodes: Map<string, ReadEntry> = new Map();
  private readonly maxEntries: number;

  constructor(maxEntries: number = FILE_READ_CACHE_MAX_ENTRIES) {
    this.maxEntries =
      Number.isInteger(maxEntries) && maxEntries >= 1
        ? maxEntries
        : FILE_READ_CACHE_MAX_ENTRIES;
  }

  /** Configured entry cap (oldest evicted first). */
  getMaxEntries(): number {
    return this.maxEntries;
  }

  /**
   * Mark a node as read
   * @param nodeId - Node identifier
   * @param content - Node content (for hash generation)
   * @param cacheable - Whether the full content was read (vs partial)
   */
  markAsRead(nodeId: string, content: string, cacheable: boolean = true): void {
    // Refresh recency on re-mark (delete + re-insert moves to most-recent).
    if (this.readNodes.has(nodeId)) this.readNodes.delete(nodeId);
    this.readNodes.set(nodeId, {
      lastReadAt: Date.now(),
      wasCacheable: cacheable,
      contentHash: this.hashContent(content)
    });
    this.evictOldest();
  }

  /** Evicts least-recently-used entries while over the cap. */
  private evictOldest(): void {
    while (this.readNodes.size > this.maxEntries) {
      const oldest = this.readNodes.keys().next().value as string | undefined;
      if (oldest === undefined) break;
      this.readNodes.delete(oldest);
    }
  }

  /** Promotes a live entry to most-recently-used. */
  private touch(nodeId: string): void {
    const entry = this.readNodes.get(nodeId);
    if (!entry) return;
    this.readNodes.delete(nodeId);
    this.readNodes.set(nodeId, entry);
  }

  /**
   * Check if a node has been read in this session
   * @param nodeId - Node identifier
   * @returns true if node was read
   */
  hasBeenRead(nodeId: string): boolean {
    const hit = this.readNodes.has(nodeId);
    if (hit) this.touch(nodeId);
    return hit;
  }

  /**
   * Check if a node was fully read (not partial/truncated)
   * @param nodeId - Node identifier
   * @returns true if node was fully read
   */
  wasFullyRead(nodeId: string): boolean {
    const entry = this.readNodes.get(nodeId);
    if (entry) this.touch(nodeId);
    return entry?.wasCacheable === true;
  }

  /**
   * Get read entry for a node
   * @param nodeId - Node identifier
   * @returns Read entry or undefined
   */
  getEntry(nodeId: string): ReadEntry | undefined {
    const entry = this.readNodes.get(nodeId);
    if (entry) this.touch(nodeId);
    return entry ? { ...entry } : undefined;
  }

  /**
   * Clear all read history (e.g., when document changes)
   */
  clear(): void {
    this.readNodes.clear();
  }

  /**
   * Get count of read nodes
   */
  getReadCount(): number {
    return this.readNodes.size;
  }

  /**
   * Get all read node IDs
   */
  getReadNodeIds(): string[] {
    return Array.from(this.readNodes.keys());
  }

  /**
   * Content identity for change detection (sha256 hex, stable across runs).
   * @param content - Content to hash
   * @returns Hash string
   */
  private hashContent(content: string): string {
    return createHash('sha256').update(content ?? '', 'utf8').digest('hex');
  }
}
