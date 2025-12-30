import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { readZipEntry } from '@zip/reader';
import { writeXlsx } from './writer';
import { bytesToString } from '../adapters/common';

describe('Column Widths', () => {
  const testFile = 'test-column-widths.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should set default column width in sheetFormatPr', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          defaultColumnWidth: 15,
          rows: (async function* () {
            yield row([cell('Test'), cell('Data')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      expect(xml).toContain('<sheetFormatPr defaultColWidth="15"/>');
    }
  });

  test('should set column width for specific column', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          columnWidths: [
            { columnIndex: 0, width: 20 },
            { columnIndex: 2, width: 25 },
          ],
          rows: (async function* () {
            yield row([cell('A'), cell('B'), cell('C')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      expect(xml).toContain('<col min="1" max="1" width="20" customWidth="1"/>');
      expect(xml).toContain('<col min="3" max="3" width="25" customWidth="1"/>');
    }
  });

  test('should set column width for column range', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          columnWidths: [
            { columnRange: { from: 0, to: 2 }, width: 18 },
          ],
          rows: (async function* () {
            yield row([cell('A'), cell('B'), cell('C'), cell('D')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      expect(xml).toContain('<col min="1" max="3" width="18" customWidth="1"/>');
    }
  });

  test('should auto-detect column widths when enabled', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          autoDetectColumnWidth: true,
          rows: (async function* () {
            yield row([cell('Short'), cell('This is a much longer string'), cell('Medium')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      // Should have cols element with auto-detected widths
      expect(xml).toContain('<cols>');
      // Second column should have larger width than first
      const colMatches = xml.match(/<col min="(\d+)" max="(\d+)" width="([\d.]+)"/g);
      expect(colMatches).toBeTruthy();
      if (colMatches) {
        expect(colMatches.length).toBeGreaterThan(0);
      }
    }
  });

  test('should override default width with specific column width', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          defaultColumnWidth: 10,
          columnWidths: [
            { columnIndex: 1, width: 30 },
          ],
          rows: (async function* () {
            yield row([cell('A'), cell('B'), cell('C')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      // Should have default width in sheetFormatPr
      expect(xml).toContain('<sheetFormatPr defaultColWidth="10"/>');
      // Should have specific width for column 2
      expect(xml).toContain('<col min="2" max="2" width="30" customWidth="1"/>');
    }
  });

  test('should auto-detect specific columns when autoDetect is set in columnWidths', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          columnWidths: [
            { columnIndex: 0, width: 15 }, // Explicit width
            { columnIndex: 1, autoDetect: true }, // Auto-detect
          ],
          rows: (async function* () {
            yield row([cell('Fixed'), cell('This is a very long auto-detected string')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      // Should have explicit width for column 1
      expect(xml).toContain('<col min="1" max="1" width="15" customWidth="1"/>');
      // Should have auto-detected width for column 2 (should be larger)
      expect(xml).toMatch(/<col min="2" max="2" width="[\d.]+" customWidth="1"\/>/);
    }
  });

  test('should auto-detect column widths for column range', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          columnWidths: [
            { columnRange: { from: 1, to: 2 }, autoDetect: true }, // Auto-detect columns B and C
          ],
          rows: (async function* () {
            yield row([cell('A'), cell('Short'), cell('This is a much longer string for column C')]);
            yield row([cell('A2'), cell('Medium'), cell('Another long string')]);
          })(),
        },
      ],
    });

    const { openZip } = await import('@zip/reader');
    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const zipFile = await openZip(buffer);
    const sheetEntry = zipFile.entries.find((e) => e.fileName === 'xl/worksheets/sheet1.xml');
    expect(sheetEntry).toBeDefined();

    if (sheetEntry) {
      const xmlBytes: Uint8Array[] = [];
      for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
        xmlBytes.push(chunk);
      }
      const xml = await bytesToString(async function* () {
        for (const chunk of xmlBytes) yield chunk;
      }());

      // Should have auto-detected width for column 2 (B)
      expect(xml).toMatch(/<col min="2" max="2" width="[\d.]+" customWidth="1"\/>/);
      // Should have auto-detected width for column 3 (C) - should be larger than column 2
      expect(xml).toMatch(/<col min="3" max="3" width="[\d.]+" customWidth="1"\/>/);

      // Extract widths to verify column 3 is wider than column 2
      const col2Match = xml.match(/<col min="2" max="2" width="([\d.]+)"/);
      const col3Match = xml.match(/<col min="3" max="3" width="([\d.]+)"/);
      expect(col2Match).toBeTruthy();
      expect(col3Match).toBeTruthy();
      if (col2Match && col3Match) {
        const width2 = parseFloat(col2Match[1]!);
        const width3 = parseFloat(col3Match[1]!);
        expect(width3).toBeGreaterThan(width2);
      }
    }
  });
});
