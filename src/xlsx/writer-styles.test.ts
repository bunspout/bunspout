/* eslint-disable @typescript-eslint/no-explicit-any */
// noinspection HtmlDeprecatedTag,XmlDeprecatedElement
// noinspection XmlDeprecatedElement
import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { openZip, readZipEntry } from '@zip/reader';
import { readXlsx } from './reader';
import { writeXlsx } from './writer';
import { bytesToString } from '../adapters';

describe('XLSXWriter - Styles', () => {
  const testFile = 'test-styles-output.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  describe('Font Styles', () => {
    test('should write cells with bold font style', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Bold Text'), style: { font: { bold: true } } }]);
          })(),
        }],
      });

      // Verify styles.xml exists
      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<b/>');
      }
    });

    test('should write cells with italic font style', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Italic Text'), style: { font: { italic: true } } }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<i/>');
      }
    });

    test('should write cells with underline font style', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Underlined'), style: { font: { underline: true } } }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<u/>');
      }
    });

    test('should write cells with strikethrough font style', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Strikethrough'), style: { font: { strikethrough: true } } }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<strike/>');
      }
    });

    test('should write cells with custom font size', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Large Text'), style: { font: { fontSize: 16 } } }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<sz val="16"/>');
      }
    });

    test('should write cells with custom font color', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Blue Text'), style: { font: { fontColor: 'FF0000FF' } } }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<color rgb="FF0000FF"/>');
      }
    });

    test('should write cells with custom font name', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{ ...cell('Custom Font'), style: { font: { fontName: 'Times New Roman' } } }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<name val="Times New Roman"/>');
      }
    });

    test('should combine multiple font properties', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          rows: (async function* () {
            yield row([{
              ...cell('Styled Text'),
              style: {
                font: {
                  bold: true,
                  italic: true,
                  fontSize: 14,
                  fontColor: 'FFFF0000',
                },
              },
            }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<b/>');
        expect(xml).toContain('<i/>');
        expect(xml).toContain('<sz val="14"/>');
        expect(xml).toContain('<color rgb="FFFF0000"/>');
      }
    });
  });

  describe('Style Deduplication', () => {
    test('should deduplicate identical styles across cells', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Dedup',
          rows: (async function* () {
            const boldStyle = { font: { bold: true, fontSize: 14 } };
            yield row([{ ...cell('Cell1'), style: boldStyle }]);
            yield row([{ ...cell('Cell2'), style: boldStyle }]);
            yield row([{ ...cell('Cell3'), style: boldStyle }]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      if (stylesEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(stylesEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Should have only one custom font (plus default)
        const fontMatches = xml.match(/<font>/g);
        expect(fontMatches?.length).toBe(2); // Default + one custom
      }
    });
  });

  describe('Empty Cells with Styles', () => {
    test('should apply styles to empty cells', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Empty',
          rows: (async function* () {
            yield row([{ ...cell(''), style: { font: { bold: true } } }]);
          })(),
        }],
      });

      // Verify styles.xml contains the style
      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      // Verify sheet XML contains the cell with style
      const sheetEntry = zipFile.entries.find(e => e.fileName === 'xl/worksheets/sheet1.xml');
      expect(sheetEntry).toBeDefined();

      if (sheetEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Should have cell with style attribute (even though value is empty)
        expect(xml).toMatch(/<c r="A1" s="\d+"/);
        expect(xml).toContain('t="inlineStr"');
      }

      // Verify file can be read back (with skipEmptyRows: false to see the empty cell)
      const workbook = await readXlsx(testFile, { skipEmptyRows: false });
      const sheet = workbook.sheet('Empty');
      const rows: any[] = [];
      for await (const row of sheet.rows()) {
        rows.push(row);
      }
      expect(rows).toHaveLength(1);
      // Cell should exist even if value is empty
      expect(rows[0]?.cells).toHaveLength(1);
    });
  });

  describe('No Styles', () => {
    test('should not create styles.xml when no styles are used', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'NoStyles',
          rows: (async function* () {
            yield row([cell('Plain Text')]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeUndefined();
    });

    test('should not emit s attribute for cells without styles', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'NoStyles',
          rows: (async function* () {
            yield row([cell('Plain')]);
          })(),
        }],
      });

      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const sheetEntry = zipFile.entries.find(e => e.fileName === 'xl/worksheets/sheet1.xml');
      expect(sheetEntry).toBeDefined();

      if (sheetEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Should not have s="0" attribute (Excel defaults to style 0)
        expect(xml).not.toContain(' s="0"');
        // Should have cell reference
        expect(xml).toContain('r="A1"');
      }
    });
  });

  describe('Styles with Column Width Tracking', () => {
    test('should apply styles when column width auto-detection is enabled', async () => {
      await writeXlsx(testFile, {
        sheets: [{
          name: 'Styled',
          autoDetectColumnWidth: true, // This triggers the slow path
          rows: (async function* () {
            yield row([
              { ...cell('Bold Text'), style: { font: { bold: true } } },
              { ...cell('Italic Text'), style: { font: { italic: true } } },
            ]);
            yield row([
              { ...cell('Red Text'), style: { font: { fontColor: 'FFFF0000' } } },
              { ...cell('Big Text'), style: { font: { fontSize: 18 } } },
            ]);
          })(),
        }],
      });

      // Verify styles.xml exists
      const zipFile = await openZip(Buffer.from(await Bun.file(testFile).arrayBuffer()));
      const stylesEntry = zipFile.entries.find(e => e.fileName === 'xl/styles.xml');
      expect(stylesEntry).toBeDefined();

      // Verify sheet XML contains cells with style attributes
      const sheetEntry = zipFile.entries.find(e => e.fileName === 'xl/worksheets/sheet1.xml');
      expect(sheetEntry).toBeDefined();

      if (sheetEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(sheetEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Should have cells with style attributes (s="...")
        expect(xml).toMatch(/<c r="A1" s="\d+"/);
        expect(xml).toMatch(/<c r="B1" s="\d+"/);
        expect(xml).toMatch(/<c r="A2" s="\d+"/);
        expect(xml).toMatch(/<c r="B2" s="\d+"/);
      }
    });
  });
});
