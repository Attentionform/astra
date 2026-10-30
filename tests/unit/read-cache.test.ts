/**
 * Unit tests for FileReadCache (Prior Read Enforcement)
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { createHash } from 'crypto';
import { FileReadCache, FILE_READ_CACHE_MAX_ENTRIES } from '@attentionform/astra-engine';

describe('FileReadCache', () => {
  let cache: FileReadCache;

  beforeEach(() => {
    cache = new FileReadCache();
  });

  describe('markAsRead', () => {
    it('should mark a node as read', () => {
      cache.markAsRead('node1', '<w:p>Test</w:p>', true);
      expect(cache.hasBeenRead('node1')).toBe(true);
    });

    it('should store read metadata', () => {
      cache.markAsRead('node1', '<w:p>Test</w:p>', true);
      const entry = cache.getEntry('node1');
      
      expect(entry).toBeDefined();
      expect(entry?.wasCacheable).toBe(true);
      expect(entry?.lastReadAt).toBeGreaterThan(0);
      expect(entry?.contentHash).toBeDefined();
    });

    it('should handle partial reads', () => {
      cache.markAsRead('node1', '<w:p>Test</w:p>', false);
      expect(cache.wasFullyRead('node1')).toBe(false);
    });
  });

  describe('hasBeenRead', () => {
    it('should return false for unread nodes', () => {
      expect(cache.hasBeenRead('node1')).toBe(false);
    });

    it('should return true for read nodes', () => {
      cache.markAsRead('node1', '<w:p>Test</w:p>', true);
      expect(cache.hasBeenRead('node1')).toBe(true);
    });
  });

  describe('wasFullyRead', () => {
    it('should return true for fully read nodes', () => {
      cache.markAsRead('node1', '<w:p>Test</w:p>', true);
      expect(cache.wasFullyRead('node1')).toBe(true);
    });

    it('should return false for partially read nodes', () => {
      cache.markAsRead('node1', '<w:p>Test</w:p>', false);
      expect(cache.wasFullyRead('node1')).toBe(false);
    });

    it('should return false for unread nodes', () => {
      expect(cache.wasFullyRead('node1')).toBe(false);
    });
  });

  describe('clear', () => {
    it('should clear all read history', () => {
      cache.markAsRead('node1', '<w:p>Test1</w:p>', true);
      cache.markAsRead('node2', '<w:p>Test2</w:p>', true);
      
      expect(cache.getReadCount()).toBe(2);
      
      cache.clear();
      
      expect(cache.getReadCount()).toBe(0);
      expect(cache.hasBeenRead('node1')).toBe(false);
      expect(cache.hasBeenRead('node2')).toBe(false);
    });
  });

  describe('getReadCount', () => {
    it('should return 0 for empty cache', () => {
      expect(cache.getReadCount()).toBe(0);
    });

    it('should return correct count', () => {
      cache.markAsRead('node1', '<w:p>Test1</w:p>', true);
      cache.markAsRead('node2', '<w:p>Test2</w:p>', true);
      cache.markAsRead('node3', '<w:p>Test3</w:p>', true);
      
      expect(cache.getReadCount()).toBe(3);
    });
  });

  describe('getReadNodeIds', () => {
    it('should return empty array for empty cache', () => {
      expect(cache.getReadNodeIds()).toEqual([]);
    });

    it('should return all read node IDs', () => {
      cache.markAsRead('node1', '<w:p>Test1</w:p>', true);
      cache.markAsRead('node2', '<w:p>Test2</w:p>', true);
      
      const ids = cache.getReadNodeIds();
      expect(ids).toContain('node1');
      expect(ids).toContain('node2');
      expect(ids.length).toBe(2);
    });
  });

  describe('content hashing', () => {
    it('should generate different hashes for different content', () => {
      cache.markAsRead('node1', '<w:p>Test1</w:p>', true);
      cache.markAsRead('node2', '<w:p>Test2</w:p>', true);
      
      const entry1 = cache.getEntry('node1');
      const entry2 = cache.getEntry('node2');
      
      expect(entry1?.contentHash).not.toBe(entry2?.contentHash);
    });

    it('should generate same hash for same content', () => {
      const content = '<w:p>Test</w:p>';
      cache.markAsRead('node1', content, true);
      cache.markAsRead('node2', content, true);
      
      const entry1 = cache.getEntry('node1');
      const entry2 = cache.getEntry('node2');
      
      expect(entry1?.contentHash).toBe(entry2?.contentHash);
    });

    it('should use sha256 hex for contentHash', () => {
      const content = '<w:p>Stable</w:p>';
      cache.markAsRead('node1', content, true);
      const entry = cache.getEntry('node1');
      const expected = createHash('sha256').update(content, 'utf8').digest('hex');
      expect(entry?.contentHash).toBe(expected);
      expect(entry?.contentHash).toMatch(/^[0-9a-f]{64}$/);
    });

    it('should keep a stable hash across separate marks', () => {
      const content = '<w:p>Pinned</w:p>';
      cache.markAsRead('a', content, true);
      const first = cache.getEntry('a')?.contentHash;
      cache.markAsRead('b', content, true);
      const second = cache.getEntry('b')?.contentHash;
      expect(first).toBe(second);
      expect(first).toBe(createHash('sha256').update(content, 'utf8').digest('hex'));
    });
  });

  describe('bounded LRU', () => {
    it('defaults to a thousands-scale cap', () => {
      expect(FILE_READ_CACHE_MAX_ENTRIES).toBeGreaterThanOrEqual(1000);
      expect(cache.getMaxEntries()).toBe(FILE_READ_CACHE_MAX_ENTRIES);
    });

    it('evicts the oldest entry past the cap', () => {
      const tiny = new FileReadCache(3);
      tiny.markAsRead('n1', 'one', true);
      tiny.markAsRead('n2', 'two', true);
      tiny.markAsRead('n3', 'three', true);
      expect(tiny.getReadCount()).toBe(3);
      tiny.markAsRead('n4', 'four', true);
      expect(tiny.getReadCount()).toBe(3);
      expect(tiny.hasBeenRead('n1')).toBe(false);
      expect(tiny.hasBeenRead('n2')).toBe(true);
      expect(tiny.hasBeenRead('n3')).toBe(true);
      expect(tiny.hasBeenRead('n4')).toBe(true);
    });

    it('recent use protects an entry from eviction', () => {
      const tiny = new FileReadCache(3);
      tiny.markAsRead('n1', 'one', true);
      tiny.markAsRead('n2', 'two', true);
      tiny.markAsRead('n3', 'three', true);
      // Touch n1 so n2 becomes least-recently-used.
      expect(tiny.hasBeenRead('n1')).toBe(true);
      tiny.markAsRead('n4', 'four', true);
      expect(tiny.hasBeenRead('n2')).toBe(false);
      expect(tiny.hasBeenRead('n1')).toBe(true);
      expect(tiny.hasBeenRead('n3')).toBe(true);
      expect(tiny.hasBeenRead('n4')).toBe(true);
    });

    it('recent entries stay intact under sustained overflow', () => {
      const tiny = new FileReadCache(5);
      for (let i = 0; i < 20; i++) tiny.markAsRead(`k${i}`, `v${i}`, true);
      expect(tiny.getReadCount()).toBe(5);
      for (let i = 15; i < 20; i++) expect(tiny.hasBeenRead(`k${i}`)).toBe(true);
      for (let i = 0; i < 15; i++) expect(tiny.hasBeenRead(`k${i}`)).toBe(false);
    });
  });
});
