/**
 * Unit tests for Grep functionality
 */

import { describe, it, expect } from 'vitest';
import { grepNodes } from '@attentionform/astra-engine';

describe('grepNodes', () => {
  const testNodes = [
    { id: 'p_001', text: 'This is the introduction section.', preview: 'Introduction' },
    { id: 'p_002', text: 'Chapter 1 discusses the methodology.', preview: 'Chapter 1' },
    { id: 'p_003', text: 'Chapter 2 presents the results.', preview: 'Chapter 2' },
    { id: 'p_004', text: 'The conclusion summarizes findings.', preview: 'Conclusion' },
  ];

  describe('literal text search', () => {
    it('should find exact text matches', () => {
      const results = grepNodes(testNodes, 'introduction');
      
      expect(results.length).toBe(1);
      expect(results[0].id).toBe('p_001');
    });

    it('should be case-insensitive by default', () => {
      const results = grepNodes(testNodes, 'chapter');
      
      expect(results.length).toBe(2);
      expect(results.map(r => r.id)).toContain('p_002');
      expect(results.map(r => r.id)).toContain('p_003');
    });

    it('should return empty array for no matches', () => {
      const results = grepNodes(testNodes, 'nonexistent');
      
      expect(results.length).toBe(0);
    });

    it('should include match context', () => {
      const results = grepNodes(testNodes, 'methodology');
      
      expect(results.length).toBe(1);
      expect(results[0].matchContext).toContain('methodology');
    });
  });

  describe('case-sensitive search', () => {
    it('should respect case when caseSensitive is true', () => {
      const results = grepNodes(testNodes, 'chapter', { caseSensitive: true });
      
      expect(results.length).toBe(0); // 'chapter' vs 'Chapter'
    });

    it('should find matches with correct case', () => {
      const results = grepNodes(testNodes, 'Chapter', { caseSensitive: true });
      
      expect(results.length).toBe(2);
    });
  });

  describe('regex search', () => {
    it('should support regex patterns', () => {
      const results = grepNodes(testNodes, 'Chapter \\d', { regex: true });
      
      expect(results.length).toBe(2);
      expect(results.map(r => r.id)).toContain('p_002');
      expect(results.map(r => r.id)).toContain('p_003');
    });

    it('should handle word boundaries', () => {
      const results = grepNodes(testNodes, '\\bthe\\b', { regex: true });
      
      expect(results.length).toBeGreaterThan(0);
    });

    it('should handle invalid regex gracefully', () => {
      const results = grepNodes(testNodes, '[invalid(', { regex: true });
      
      expect(results.length).toBe(0);
    });
  });

  describe('match context extraction', () => {
    it('should include surrounding text', () => {
      const results = grepNodes(testNodes, 'methodology');
      
      expect(results[0].matchContext).toContain('discusses');
      expect(results[0].matchContext).toContain('methodology');
    });

    it('should add ellipsis for truncated context', () => {
      const longText = 'a'.repeat(200) + 'target' + 'b'.repeat(200);
      const nodes = [{ id: 'p_001', text: longText, preview: 'Long' }];
      
      const results = grepNodes(nodes, 'target');
      
      expect(results[0].matchContext).toContain('...');
    });
  });

  describe('special characters', () => {
    it('should escape special regex characters in literal search', () => {
      const nodes = [
        { id: 'p_001', text: 'Price is $100.00', preview: 'Price' }
      ];
      
      const results = grepNodes(nodes, '$100.00');
      
      expect(results.length).toBe(1);
    });

    it('should handle parentheses', () => {
      const nodes = [
        { id: 'p_001', text: 'See (Figure 1)', preview: 'Figure' }
      ];
      
      const results = grepNodes(nodes, '(Figure 1)');
      
      expect(results.length).toBe(1);
    });
  });
});
