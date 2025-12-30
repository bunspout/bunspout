import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { readZipEntry } from '@zip/reader';
import { writeXlsx } from './writer';
import { bytesToString } from '../adapters/common';

describe('Row Heights', () => {
  const testFile = 'test-row-heights.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should set default row height in sheetFormatPr', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          defaultRowHeight: 20,
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

      expect(xml).toContain('<sheetFormatPr defaultRowHeight="20"/>');
    }
  });

  test('should set row height for row with direct height option', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            yield row([cell('A'), cell('B')], { height: 25 });
            yield row([cell('C'), cell('D')]);
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

      // First row should have height attribute
      expect(xml).toContain('<row r="1"');
      expect(xml).toMatch(/<row r="1"[^>]*ht="25" customHeight="1"/);
      // Second row should not have height attribute
      expect(xml).toContain('<row r="2"');
      expect(xml).not.toMatch(/<row r="2"[^>]*ht=/);
    }
  });

  test('should set row height for specific row via rowHeights', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rowHeights: [
            { rowIndex: 2, height: 30 },
            { rowIndex: 4, height: 35 },
          ],
          rows: (async function* () {
            yield row([cell('Row1')]);
            yield row([cell('Row2')]);
            yield row([cell('Row3')]);
            yield row([cell('Row4')]);
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

      // Row 2 should have height 30
      expect(xml).toMatch(/<row r="2"[^>]*ht="30" customHeight="1"/);
      // Row 4 should have height 35
      expect(xml).toMatch(/<row r="4"[^>]*ht="35" customHeight="1"/);
      // Rows 1 and 3 should not have height attributes
      expect(xml).toContain('<row r="1"');
      expect(xml).not.toMatch(/<row r="1"[^>]*ht=/);
      expect(xml).toContain('<row r="3"');
      expect(xml).not.toMatch(/<row r="3"[^>]*ht=/);
    }
  });

  test('should set row height for row range via rowHeights', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rowHeights: [
            { rowRange: { from: 2, to: 4 }, height: 28 },
          ],
          rows: (async function* () {
            yield row([cell('Row1')]);
            yield row([cell('Row2')]);
            yield row([cell('Row3')]);
            yield row([cell('Row4')]);
            yield row([cell('Row5')]);
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

      // Rows 2, 3, and 4 should have height 28
      expect(xml).toMatch(/<row r="2"[^>]*ht="28" customHeight="1"/);
      expect(xml).toMatch(/<row r="3"[^>]*ht="28" customHeight="1"/);
      expect(xml).toMatch(/<row r="4"[^>]*ht="28" customHeight="1"/);
      // Rows 1 and 5 should not have height attributes
      expect(xml).toContain('<row r="1"');
      expect(xml).not.toMatch(/<row r="1"[^>]*ht=/);
      expect(xml).toContain('<row r="5"');
      expect(xml).not.toMatch(/<row r="5"[^>]*ht=/);
    }
  });

  test('should prioritize direct row height over rowHeights definition', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rowHeights: [
            { rowIndex: 1, height: 20 },
          ],
          rows: (async function* () {
            yield row([cell('Row1')], { height: 40 }); // Direct height should take priority
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

      // Should use direct height (40) not rowHeights definition (20)
      expect(xml).toMatch(/<row r="1"[^>]*ht="40" customHeight="1"/);
      expect(xml).not.toMatch(/<row r="1"[^>]*ht="20"/);
    }
  });

  test('should combine default row height with explicit row heights', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          defaultRowHeight: 18,
          rowHeights: [
            { rowIndex: 2, height: 30 },
          ],
          rows: (async function* () {
            yield row([cell('Row1')]); // Should use default (18) - but default doesn't add ht attribute
            yield row([cell('Row2')]); // Should use explicit (30)
            yield row([cell('Row3')], { height: 25 }); // Should use direct (25)
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

      // Should have defaultRowHeight in sheetFormatPr (may be combined with defaultColWidth if present)
      expect(xml).toMatch(/<sheetFormatPr[^>]*defaultRowHeight="18"/);
      // Row 1 should not have height attribute (uses default, which doesn't add ht attribute)
      expect(xml).toContain('<row r="1"');
      expect(xml).not.toMatch(/<row r="1"[^>]*ht=/);
      // Row 2 should have explicit height from rowHeights
      expect(xml).toMatch(/<row r="2"[^>]*ht="30" customHeight="1"/);
      // Row 3 should have direct height
      expect(xml).toMatch(/<row r="3"[^>]*ht="25" customHeight="1"/);
    }
  });
});
