import { describe, test, expect } from 'bun:test';
import { SharedStringsTable } from './shared-strings';

describe('SharedStringsTable', () => {
  test('should add strings and return indices', () => {
    const table = new SharedStringsTable();
    const index1 = table.addString('Hello');
    const index2 = table.addString('World');
    const index3 = table.addString('Hello'); // Duplicate

    expect(index1).toBe(0);
    expect(index2).toBe(1);
    expect(index3).toBe(0); // Should return same index for duplicate
    expect(table.getUniqueCount()).toBe(2);
  });

  test('should get index of added string', () => {
    const table = new SharedStringsTable();
    table.addString('Test');
    expect(table.getIndex('Test')).toBe(0);
  });

  test('should track total count and unique count separately', () => {
    const table = new SharedStringsTable();
    table.addString('A');
    table.addString('B');
    table.addString('A'); // Duplicate
    table.addString('A'); // Duplicate again

    expect(table.getUniqueCount()).toBe(2); // A and B
    expect(table.getTotalCount()).toBe(4); // 4 total references
  });

  test('should throw error for string not in table', () => {
    const table = new SharedStringsTable();
    expect(() => table.getIndex('NotInTable')).toThrow();
  });

  test('should generate XML for empty table', () => {
    const table = new SharedStringsTable();
    const xml = table.generateXml();
    expect(xml).toContain('<sst');
    expect(xml).toContain('count="0"');
    expect(xml).toContain('uniqueCount="0"');
  });

  test('should generate XML with strings', () => {
    const table = new SharedStringsTable();
    table.addString('Hello');
    table.addString('World');
    const xml = table.generateXml();

    expect(xml).toContain('count="2"');
    expect(xml).toContain('uniqueCount="2"');
    expect(xml).toContain('<si><t>Hello</t></si>');
    expect(xml).toContain('<si><t>World</t></si>');
  });

  test('should escape XML special characters in strings', () => {
    const table = new SharedStringsTable();
    table.addString('<test>&"value"</test>');
    const xml = table.generateXml();

    expect(xml).toContain('&lt;test&gt;&amp;&quot;value&quot;&lt;/test&gt;');
    expect(xml).not.toContain('<test>&"value"</test>');
  });

  test('should handle empty strings', () => {
    const table = new SharedStringsTable();
    table.addString('');
    const xml = table.generateXml();

    expect(xml).toContain('<si><t></t></si>');
  });
});
