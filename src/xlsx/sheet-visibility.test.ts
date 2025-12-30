/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { readXlsx } from './reader';
import { generateWorkbook } from './structure';
import { writeXlsx } from './writer';

describe('Sheet Visibility', () => {
  const testFile = 'test-visibility.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  describe('generateWorkbook', () => {
    test('should include state="hidden" attribute for hidden sheets', () => {
      const result = generateWorkbook([
        { name: 'Visible Sheet', id: 1 },
        { name: 'Hidden Sheet', id: 2, hidden: true },
        { name: 'Another Visible', id: 3 },
      ]);

      expect(result).toContain('<sheet name="Visible Sheet" sheetId="1" r:id="rId1"/>');
      expect(result).toContain('<sheet name="Hidden Sheet" sheetId="2" r:id="rId2" state="hidden"/>');
      expect(result).toContain('<sheet name="Another Visible" sheetId="3" r:id="rId3"/>');

      // Verify visible sheets don't have state attribute
      expect(result).toMatch(/<sheet name="Visible Sheet"[^>]*\/>/);
      expect(result).toMatch(/<sheet name="Another Visible"[^>]*\/>/);

      // Verify hidden sheet has state="hidden"
      expect(result).toMatch(/<sheet name="Hidden Sheet"[^>]*state="hidden"/);

      // Count hidden attributes - should be exactly 1
      const hiddenMatches = result.match(/state="hidden"/g);
      expect(hiddenMatches).toHaveLength(1);
    });

    test('should not include state attribute for visible sheets', () => {
      const result = generateWorkbook([
        { name: 'Sheet1', id: 1 },
        { name: 'Sheet2', id: 2 },
      ]);

      expect(result).not.toContain('state=');
    });

    test('should handle all sheets hidden', () => {
      const result = generateWorkbook([
        { name: 'Hidden1', id: 1, hidden: true },
        { name: 'Hidden2', id: 2, hidden: true },
      ]);

      expect(result).toContain('state="hidden"');
      const hiddenMatches = result.match(/state="hidden"/g);
      expect(hiddenMatches).toHaveLength(2);
    });
  });

  describe('writeXlsx with hidden sheets', () => {
    test('should write workbook with hidden sheet', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Visible',
            rows: (async function* () {
              yield row([cell('Visible Data')]);
            })(),
          },
          {
            name: 'Hidden',
            hidden: true,
            rows: (async function* () {
              yield row([cell('Hidden Data')]);
            })(),
          },
        ],
      });

      const file = Bun.file(testFile);
      expect(await file.exists()).toBe(true);

      // Verify workbook.xml contains hidden state
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const workbookEntry = zipFile.entries.find((e) => e.fileName === 'xl/workbook.xml');
      expect(workbookEntry).toBeDefined();

      if (workbookEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(workbookEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<sheet name="Visible"');
        expect(xml).toContain('<sheet name="Hidden"');
        expect(xml).toContain('state="hidden"');
        // Verify hidden attribute is on the Hidden sheet, not Visible
        const hiddenSheetMatch = xml.match(/<sheet name="Hidden"[^>]*state="hidden"/);
        expect(hiddenSheetMatch).toBeTruthy();
      }
    });

    test('should write workbook with multiple hidden and visible sheets', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Sheet1',
            rows: (async function* () {
              yield row([cell('Data1')]);
            })(),
          },
          {
            name: 'Hidden1',
            hidden: true,
            rows: (async function* () {
              yield row([cell('Hidden1')]);
            })(),
          },
          {
            name: 'Sheet2',
            rows: (async function* () {
              yield row([cell('Data2')]);
            })(),
          },
          {
            name: 'Hidden2',
            hidden: true,
            rows: (async function* () {
              yield row([cell('Hidden2')]);
            })(),
          },
        ],
      });

      // Verify workbook.xml
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const workbookEntry = zipFile.entries.find((e) => e.fileName === 'xl/workbook.xml');
      expect(workbookEntry).toBeDefined();

      if (workbookEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(workbookEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Verify visible sheets don't have state attribute
        expect(xml).toMatch(/<sheet name="Sheet1"[^>]*\/>/);
        expect(xml).toMatch(/<sheet name="Sheet2"[^>]*\/>/);

        // Verify hidden sheets have state="hidden"
        expect(xml).toMatch(/<sheet name="Hidden1"[^>]*state="hidden"/);
        expect(xml).toMatch(/<sheet name="Hidden2"[^>]*state="hidden"/);

        // Count hidden attributes
        const hiddenMatches = xml.match(/state="hidden"/g);
        expect(hiddenMatches).toHaveLength(2);
      }
    });

    test('should read workbook with hidden sheets correctly', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Visible',
            rows: (async function* () {
              yield row([cell('Visible')]);
            })(),
          },
          {
            name: 'Hidden',
            hidden: true,
            rows: (async function* () {
              yield row([cell('Hidden')]);
            })(),
          },
        ],
      });

      // Verify we can read it back (hidden sheets are still accessible via API)
      const workbook = await readXlsx(testFile);
      expect(workbook.sheets()).toHaveLength(2);

      // Both sheets should be accessible
      const visibleSheet = workbook.sheet('Visible');
      expect(visibleSheet.name).toBe('Visible');

      const hiddenSheet = workbook.sheet('Hidden');
      expect(hiddenSheet.name).toBe('Hidden');

      // Verify data is still readable
      const visibleRows: any[] = [];
      for await (const row of visibleSheet.rows()) {
        visibleRows.push(row);
      }
      expect(visibleRows[0]?.cells[0]?.value).toBe('Visible');

      const hiddenRows: any[] = [];
      for await (const row of hiddenSheet.rows()) {
        hiddenRows.push(row);
      }
      expect(hiddenRows[0]?.cells[0]?.value).toBe('Hidden');
    });
  });
});
