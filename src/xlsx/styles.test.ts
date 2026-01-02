// noinspection HtmlDeprecatedTag,XmlDeprecatedElement
// noinspection XmlDeprecatedElement
import { describe, test, expect } from 'bun:test';
import { StyleRegistry } from './styles';
import type { Style } from '../types';

describe('StyleRegistry', () => {
  describe('addStyle', () => {
    test('should register a style and return cellXfs index >= 1', () => {
      const registry = new StyleRegistry();
      const style: Style = {
        font: { bold: true },
      };
      const index = registry.addStyle(style);
      expect(index).toBeGreaterThanOrEqual(1);
      expect(registry.getCount()).toBe(1);
    });

    test('should deduplicate identical styles', () => {
      const registry = new StyleRegistry();
      const style1: Style = { font: { bold: true, fontSize: 14 } };
      const style2: Style = { font: { bold: true, fontSize: 14 } };

      const index1 = registry.addStyle(style1);
      const index2 = registry.addStyle(style2);

      expect(index1).toBe(index2);
      expect(registry.getCount()).toBe(1);
    });

    test('should assign different indices to different styles', () => {
      const registry = new StyleRegistry();
      const style1: Style = { font: { bold: true } };
      const style2: Style = { font: { italic: true } };

      const index1 = registry.addStyle(style1);
      const index2 = registry.addStyle(style2);

      expect(index1).not.toBe(index2);
      expect(registry.getCount()).toBe(2);
    });

    test('should handle font styles: bold', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { bold: true } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<b/>');
    });

    test('should handle font styles: italic', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { italic: true } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<i/>');
    });

    test('should handle font styles: underline', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { underline: true } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<u/>');
    });

    test('should handle font styles: strikethrough', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { strikethrough: true } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<strike/>');
    });

    test('should handle font size', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { fontSize: 16 } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<sz val="16"/>');
    });

    test('should handle font color', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { fontColor: 'FF0000FF' } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<color rgb="FF0000FF"/>');
    });

    test('should handle font name', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { fontName: 'Times New Roman' } };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<name val="Times New Roman"/>');
    });

    test('should combine multiple font properties', () => {
      const registry = new StyleRegistry();
      const style: Style = {
        font: {
          bold: true,
          italic: true,
          fontSize: 14,
          fontColor: 'FFFF0000',
          fontName: 'Courier New',
        },
      };
      registry.addStyle(style);
      const xml = registry.generateXml();
      expect(xml).toContain('<b/>');
      expect(xml).toContain('<i/>');
      expect(xml).toContain('<sz val="14"/>');
      expect(xml).toContain('<color rgb="FFFF0000"/>');
      expect(xml).toContain('<name val="Courier New"/>');
    });

    test('should validate font color format', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { fontColor: 'INVALID' } };
      expect(() => registry.addStyle(style)).toThrow();
    });

    test('should validate font size range', () => {
      const registry = new StyleRegistry();
      const style1: Style = { font: { fontSize: 0 } };
      expect(() => registry.addStyle(style1)).toThrow();

      const style2: Style = { font: { fontSize: 500 } };
      expect(() => registry.addStyle(style2)).toThrow();
    });
  });

  describe('getCellXfIndex', () => {
    test('should return cellXfs index for registered style', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { bold: true } };
      const addIndex = registry.addStyle(style);
      const getIndex = registry.getCellXfIndex(style);
      expect(getIndex).toBe(addIndex);
      expect(getIndex).toBeGreaterThanOrEqual(1);
    });

    test('should throw if style not registered', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { bold: true } };
      expect(() => registry.getCellXfIndex(style)).toThrow();
    });
  });

  describe('generateXml', () => {
    test('should generate minimal styles.xml when no styles registered', () => {
      const registry = new StyleRegistry();
      const xml = registry.generateXml();
      expect(xml).toContain('<styleSheet');
      expect(xml).toContain('<fonts count="1">');
      expect(xml).toContain('<cellXfs count="1">');
    });

    test('should generate styles.xml with registered styles', () => {
      const registry = new StyleRegistry();
      registry.addStyle({ font: { bold: true } });
      registry.addStyle({ font: { italic: true } });
      const xml = registry.generateXml();
      expect(xml).toContain('<fonts count="');
      expect(xml).toContain('<cellXfs count="');
      expect(registry.getCount()).toBe(2);
    });

    test('should include default font at index 0', () => {
      const registry = new StyleRegistry();
      registry.addStyle({ font: { bold: true } });
      const xml = registry.generateXml();
      // Default font should be first
      expect(xml).toMatch(/<fonts count="\d+">[\s\S]*?<sz val="11"\/>/);
    });
  });

  describe('Invariants', () => {
    test('should enforce structural immutability - stored styles cannot be mutated', () => {
      const registry = new StyleRegistry();
      const style: Style = { font: { bold: true, fontSize: 14 } };
      const index = registry.addStyle(style);

      // Attempt to mutate the input style (should not affect registry)
      style.font!.bold = false;
      style.font!.fontSize = 16;

      // Registry should still return the same index for the original style
      const originalStyle: Style = { font: { bold: true, fontSize: 14 } };
      const sameIndex = registry.addStyle(originalStyle);
      expect(sameIndex).toBe(index);
      expect(registry.getCount()).toBe(1);
    });

    test('should use stable hash for deduplication', () => {
      const registry = new StyleRegistry();
      const style1: Style = { font: { bold: true, fontSize: 14, fontColor: 'FF0000FF' } };
      const style2: Style = { font: { bold: true, fontSize: 14, fontColor: 'FF0000FF' } };
      const style3: Style = { font: { bold: true, fontSize: 14, fontColor: 'FF0000FF' } };

      const index1 = registry.addStyle(style1);
      const index2 = registry.addStyle(style2);
      const index3 = registry.addStyle(style3);

      // All identical styles should get the same index
      expect(index1).toBe(index2);
      expect(index2).toBe(index3);
      expect(registry.getCount()).toBe(1);
    });

    test('should maintain deterministic ordering', () => {
      const registry1 = new StyleRegistry();
      const registry2 = new StyleRegistry();

      const styles: Style[] = [
        { font: { bold: true } },
        { font: { italic: true } },
        { font: { underline: true } },
        { font: { bold: true, italic: true } },
      ];

      // Add styles in same order to both registries
      const indices1 = styles.map(s => registry1.addStyle(s));
      const indices2 = styles.map(s => registry2.addStyle(s));

      // Indices should match (deterministic ordering)
      expect(indices1).toEqual(indices2);

      // Adding duplicates should return same indices
      const duplicateIndices1 = styles.map(s => registry1.addStyle(s));
      const duplicateIndices2 = styles.map(s => registry2.addStyle(s));

      expect(duplicateIndices1).toEqual(indices1);
      expect(duplicateIndices2).toEqual(indices2);
    });

    test('should handle edge cases in hash key generation', () => {
      const registry = new StyleRegistry();

      // Test with all properties set
      const fullStyle: Style = {
        font: {
          bold: true,
          italic: true,
          underline: true,
          strikethrough: true,
          fontSize: 20,
          fontColor: 'FFFFFFFF',
          fontName: 'Times New Roman',
        },
      };

      const index1 = registry.addStyle(fullStyle);
      const index2 = registry.addStyle(fullStyle);
      expect(index1).toBe(index2);

      // Test with minimal style (only defaults)
      const minimalStyle: Style = {};
      const index3 = registry.addStyle(minimalStyle);
      const index4 = registry.addStyle(minimalStyle);
      expect(index3).toBe(index4);

      // Different styles should get different indices
      expect(index1).not.toBe(index3);
    });

    test('should preserve ordering even when styles are added out of order', () => {
      const registry = new StyleRegistry();

      const styleA: Style = { font: { bold: true } };
      const styleB: Style = { font: { italic: true } };
      const styleC: Style = { font: { underline: true } };

      // Add A, B, C
      const indexA1 = registry.addStyle(styleA);
      const indexB1 = registry.addStyle(styleB);
      const indexC1 = registry.addStyle(styleC);

      // Add C, A, B (different order)
      const indexC2 = registry.addStyle(styleC);
      const indexA2 = registry.addStyle(styleA);
      const indexB2 = registry.addStyle(styleB);

      // First-time additions should maintain insertion order
      expect(indexA1).toBe(1);
      expect(indexB1).toBe(2);
      expect(indexC1).toBe(3);

      // Duplicates should return original indices
      expect(indexC2).toBe(indexC1);
      expect(indexA2).toBe(indexA1);
      expect(indexB2).toBe(indexB1);
    });
  });
});
