import { describe, test, expect } from 'bun:test';
import {
  generateContentTypes,
  generateRels,
  generateWorkbook,
  generateWorkbookRels,
  generateCustomProperties,
} from './structure';

describe('XLSX Structure Generators', () => {
  describe('generateContentTypes', () => {
    test('should generate Content_Types.xml for single sheet', () => {
      const result = generateContentTypes([{ id: 1 }]);
      expect(result).toContain('application/vnd.openxmlformats-package.relationships+xml');
      expect(result).toContain('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
      expect(result).toContain('/xl/worksheets/sheet1.xml');
      expect(result).toMatch(/<Override PartName="\/xl\/worksheets\/sheet1\.xml"/);
    });

    test('should include shared strings in Content_Types.xml when enabled', () => {
      const result = generateContentTypes([{ id: 1 }], true);
      expect(result).toContain('/xl/sharedStrings.xml');
      expect(result).toContain('application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
    });

    test('should include core properties in Content_Types.xml when enabled', () => {
      const result = generateContentTypes([{ id: 1 }], false, true);
      expect(result).toContain('/docProps/core.xml');
      expect(result).toContain('application/vnd.openxmlformats-package.core-properties+xml');
    });

    test('should include custom properties in Content_Types.xml when enabled', () => {
      const result = generateContentTypes([{ id: 1 }], false, false, true);
      expect(result).toContain('/docProps/custom.xml');
      expect(result).toContain('application/vnd.openxmlformats-officedocument.custom-properties+xml');
    });

    test('should generate Content_Types.xml for multiple sheets', () => {
      const result = generateContentTypes([{ id: 1 }, { id: 2 }, { id: 3 }]);
      expect(result).toContain('/xl/worksheets/sheet1.xml');
      expect(result).toContain('/xl/worksheets/sheet2.xml');
      expect(result).toContain('/xl/worksheets/sheet3.xml');
      const sheetMatches = result.match(/xl\/worksheets\/sheet\d+\.xml/g);
      expect(sheetMatches).toHaveLength(3);
    });

    test('should handle non-sequential sheet IDs', () => {
      const result = generateContentTypes([{ id: 1 }, { id: 5 }]);
      expect(result).toContain('/xl/worksheets/sheet1.xml');
      expect(result).toContain('/xl/worksheets/sheet5.xml');
    });
  });

  describe('generateRels', () => {
    test('should generate _rels/.rels file', () => {
      const result = generateRels();
      expect(result).toContain('http://schemas.openxmlformats.org/package/2006/relationships');
      expect(result).toContain('xl/workbook.xml');
      expect(result).toMatch(/<Relationship Id="rId1"/);
      expect(result).toMatch(/Target="xl\/workbook\.xml"/);
    });

    test('should include core properties relationship when enabled', () => {
      const result = generateRels(true);
      expect(result).toContain('core.xml');
      expect(result).toMatch(/<Relationship Id="rId2" Type="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships\/metadata\/core-properties" Target="docProps\/core\.xml"/);
    });

    test('should include custom properties relationship when enabled', () => {
      const result = generateRels(false, true);
      expect(result).toContain('custom.xml');
      expect(result).toMatch(/<Relationship Id="rId2" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/custom-properties" Target="docProps\/custom\.xml"/);
    });

    test('should handle core and custom properties together in _rels/.rels', () => {
      const result = generateRels(true, true);
      expect(result).toContain('core.xml');
      expect(result).toContain('custom.xml');
      expect(result).toMatch(/rId2.*core\.xml/);
      expect(result).toMatch(/rId3.*custom\.xml/);
    });
  });

  describe('generateWorkbook', () => {
    test('should generate workbook.xml with sheet names', () => {
      const result = generateWorkbook([
        { name: 'Sheet1', id: 1 },
        { name: 'Data', id: 2 },
      ]);
      expect(result).toContain('<workbook');
      expect(result).toContain('<sheets>');
      expect(result).toContain('Sheet1');
      expect(result).toContain('Data');
      expect(result).toMatch(/<sheet name="Sheet1" sheetId="1" r:id="rId1"/);
      expect(result).toMatch(/<sheet name="Data" sheetId="2" r:id="rId2"/);
    });

    test('should escape XML special characters in sheet names', () => {
      const result = generateWorkbook([{ name: 'Sheet & Data <Test>', id: 1 }]);
      expect(result).toContain('&amp;');
      expect(result).toContain('&lt;');
      expect(result).toContain('&gt;');
      expect(result).not.toContain('Sheet & Data <Test>');
    });

    test('should generate workbook.xml for single sheet', () => {
      const result = generateWorkbook([{ name: 'MySheet', id: 1 }]);
      expect(result).toContain('MySheet');
      expect(result).toMatch(/<sheet name="MySheet" sheetId="1" r:id="rId1"/);
    });
  });

  describe('generateWorkbookRels', () => {
    test('should generate workbook.xml.rels for single sheet', () => {
      const result = generateWorkbookRels([{ id: 1 }]);
      expect(result).toContain('worksheets/sheet1.xml');
      expect(result).toMatch(/<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/worksheet" Target="worksheets\/sheet1\.xml"/);
    });

    test('should include shared strings relationship when enabled', () => {
      const result = generateWorkbookRels([{ id: 1 }], true);
      expect(result).toContain('sharedStrings.xml');
      expect(result).toMatch(/<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/sharedStrings" Target="sharedStrings\.xml"/);
      expect(result).toMatch(/<Relationship Id="rId2" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/worksheet" Target="worksheets\/sheet1\.xml"/);
    });

    test('should handle shared strings and core properties together', () => {
      const result = generateWorkbookRels([{ id: 1 }], true, true);
      expect(result).toContain('sharedStrings.xml');
      expect(result).toMatch(/rId1.*sharedStrings/);
      expect(result).toMatch(/rId2.*sheet1/);
      // Note: core properties relationship goes in _rels/.rels, not workbook.xml.rels
    });

    test('should generate workbook.xml.rels for multiple sheets', () => {
      const result = generateWorkbookRels([{ id: 1 }, { id: 2 }, { id: 3 }]);
      expect(result).toContain('worksheets/sheet1.xml');
      expect(result).toContain('worksheets/sheet2.xml');
      expect(result).toContain('worksheets/sheet3.xml');
      expect(result).toMatch(/rId1.*sheet1/);
      expect(result).toMatch(/rId2.*sheet2/);
      expect(result).toMatch(/rId3.*sheet3/);
    });

    test('should handle non-sequential sheet IDs', () => {
      const result = generateWorkbookRels([{ id: 1 }, { id: 5 }]);
      expect(result).toContain('worksheets/sheet1.xml');
      expect(result).toContain('worksheets/sheet5.xml');
    });
  });

  describe('generateCustomProperties', () => {
    test('should generate custom.xml with properties', () => {
      const result = generateCustomProperties({
        'Custom1': 'Value1',
        'Custom2': 'Value2',
      });

      expect(result).toContain('<Properties');
      expect(result).toContain('fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"');
      expect(result).toContain('pid="2"');
      expect(result).toContain('pid="3"');
      expect(result).toContain('name="Custom1"');
      expect(result).toContain('name="Custom2"');
      expect(result).toContain('<vt:lpwstr>Value1</vt:lpwstr>');
      expect(result).toContain('<vt:lpwstr>Value2</vt:lpwstr>');
    });

    test('should return empty string for empty properties', () => {
      const result = generateCustomProperties({});
      expect(result).toBe('');
    });

    test('should escape XML special characters', () => {
      const result = generateCustomProperties({
        'Name & <Test>': 'Value with "quotes"',
      });

      expect(result).toContain('&amp;');
      expect(result).toContain('&lt;');
      expect(result).toContain('&gt;');
      expect(result).toContain('&quot;');
    });
  });
});
