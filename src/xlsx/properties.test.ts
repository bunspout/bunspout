import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { generateCoreProperties, generateCustomProperties } from './structure';
import { writeXlsx } from './writer';

describe('Document Properties', () => {
  const testFile = 'test-properties.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  describe('generateCoreProperties', () => {
    test('should generate core.xml with default properties', () => {
      const props = {
        title: 'Untitled Spreadsheet',
        creator: 'BunSpout',
        lastModifiedBy: 'BunSpout',
      };
      const xml = generateCoreProperties(props);

      expect(xml).toContain('<cp:coreProperties');
      expect(xml).toContain('<dc:title>Untitled Spreadsheet</dc:title>');
      expect(xml).toContain('<dc:creator>BunSpout</dc:creator>');
      expect(xml).toContain('<cp:lastModifiedBy>BunSpout</cp:lastModifiedBy>');
      expect(xml).toContain('<dcterms:created');
      expect(xml).toContain('<dcterms:modified');
      expect(xml).toContain('<cp:revision>0</cp:revision>');
    });

    test('should include all optional properties when provided', () => {
      const props = {
        title: 'My Spreadsheet',
        subject: 'Test Subject',
        creator: 'John Doe',
        lastModifiedBy: 'Jane Doe',
        keywords: 'test, spreadsheet',
        description: 'A test spreadsheet',
        category: 'Test',
        language: 'en-US',
      };
      const xml = generateCoreProperties(props);

      expect(xml).toContain('<dc:title>My Spreadsheet</dc:title>');
      expect(xml).toContain('<dc:subject>Test Subject</dc:subject>');
      expect(xml).toContain('<dc:creator>John Doe</dc:creator>');
      expect(xml).toContain('<cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>');
      expect(xml).toContain('<cp:keywords>test, spreadsheet</cp:keywords>');
      expect(xml).toContain('<dc:description>A test spreadsheet</dc:description>');
      expect(xml).toContain('<cp:category>Test</cp:category>');
      expect(xml).toContain('<dc:language>en-US</dc:language>');
    });

    test('should escape XML special characters in properties', () => {
      const props = {
        title: 'Test & <Title>',
        subject: 'Subject with "quotes"',
      };
      const xml = generateCoreProperties(props);

      expect(xml).toContain('&amp;');
      expect(xml).toContain('&lt;');
      expect(xml).toContain('&gt;');
      expect(xml).toContain('&quot;');
      expect(xml).not.toContain('Test & <Title>');
    });

    test('should omit null/undefined properties', () => {
      const props = {
        title: 'Only Title',
        subject: null,
        keywords: undefined,
      };
      const xml = generateCoreProperties(props);

      expect(xml).toContain('<dc:title>Only Title</dc:title>');
      expect(xml).not.toContain('<dc:subject>');
      expect(xml).not.toContain('<cp:keywords>');
    });

    test('should generate custom.xml with custom properties', () => {
      const customProps = {
        'Custom1': 'Value1',
        'Custom2': 'Value2',
        'Custom3': 'Value3',
      };
      const xml = generateCustomProperties(customProps);

      expect(xml).toContain('<Properties');
      expect(xml).toContain('xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"');
      expect(xml).toContain('xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"');
      expect(xml).toContain('fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"');
      expect(xml).toContain('pid="2"');
      expect(xml).toContain('pid="3"');
      expect(xml).toContain('pid="4"');
      expect(xml).toContain('name="Custom1"');
      expect(xml).toContain('name="Custom2"');
      expect(xml).toContain('name="Custom3"');
      expect(xml).toContain('<vt:lpwstr>Value1</vt:lpwstr>');
      expect(xml).toContain('<vt:lpwstr>Value2</vt:lpwstr>');
      expect(xml).toContain('<vt:lpwstr>Value3</vt:lpwstr>');
    });

    test('should return empty string for empty custom properties', () => {
      const xml = generateCustomProperties({});
      expect(xml).toBe('');
    });

    test('should escape XML special characters in custom property names and values', () => {
      const customProps = {
        'Name & <Test>': 'Value with "quotes"',
      };
      const xml = generateCustomProperties(customProps);

      expect(xml).toContain('&amp;');
      expect(xml).toContain('&lt;');
      expect(xml).toContain('&gt;');
      expect(xml).toContain('&quot;');
      expect(xml).not.toContain('Name & <Test>');
      expect(xml).not.toContain('Value with "quotes"');
    });
  });

  describe('writeXlsx with properties', () => {
    test('should write core.xml when properties are provided', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          title: 'My Workbook',
          creator: 'Test User',
        },
      });

      const file = Bun.file(testFile);
      expect(await file.exists()).toBe(true);

      // Verify core.xml exists in ZIP
      const { openZip } = await import('../zip/reader');
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).toContain('docProps/core.xml');
    });

    test('should write workbook with all properties', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          title: 'Complete Workbook',
          subject: 'Testing',
          creator: 'Creator',
          lastModifiedBy: 'Modifier',
          keywords: 'test, excel',
          description: 'A complete test',
          category: 'Test',
          language: 'en-US',
        },
      });

      // Read and verify core.xml content
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const coreEntry = zipFile.entries.find((e) => e.fileName === 'docProps/core.xml');
      expect(coreEntry).toBeDefined();

      if (coreEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(coreEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<dc:title>Complete Workbook</dc:title>');
        expect(xml).toContain('<dc:subject>Testing</dc:subject>');
        expect(xml).toContain('<dc:creator>Creator</dc:creator>');
        expect(xml).toContain('<cp:lastModifiedBy>Modifier</cp:lastModifiedBy>');
        expect(xml).toContain('<cp:keywords>test, excel</cp:keywords>');
        expect(xml).toContain('<dc:description>A complete test</dc:description>');
        expect(xml).toContain('<cp:category>Test</cp:category>');
        expect(xml).toContain('<dc:language>en-US</dc:language>');
      }
    });

    test('should not write core.xml when properties are not provided', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
      });

      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).not.toContain('docProps/core.xml');
    });

    test('should work with properties and shared strings together', async () => {
      await writeXlsx(
        testFile,
        {
          sheets: [
            {
              name: 'Data',
              rows: (async function* () {
                yield row([cell('Hello')]);
              })(),
            },
          ],
          properties: {
            title: 'With Shared Strings',
            creator: 'Test',
          },
        },
        { sharedStrings: 'shared' },
      );

      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).toContain('docProps/core.xml');
      expect(fileNames).toContain('xl/sharedStrings.xml');
    });

    test('should include core properties relationship in _rels/.rels', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          title: 'Test Workbook',
          creator: 'Test',
        },
      });

      // Verify _rels/.rels includes core properties relationship
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const relsEntry = zipFile.entries.find((e) => e.fileName === '_rels/.rels');
      expect(relsEntry).toBeDefined();

      if (relsEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(relsEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('core.xml');
        expect(xml).toMatch(/<Relationship Id="rId2" Type="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships\/metadata\/core-properties" Target="docProps\/core\.xml"/);
      }
    });

    test('should write custom.xml when customProperties are provided', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          title: 'Test',
          customProperties: {
            'Custom1': 'Value1',
            'Custom2': 'Value2',
          },
        },
      });

      // Verify custom.xml exists in ZIP
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).toContain('docProps/custom.xml');

      // Verify custom.xml content
      const customEntry = zipFile.entries.find((e) => e.fileName === 'docProps/custom.xml');
      expect(customEntry).toBeDefined();

      if (customEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(customEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('<Properties');
        expect(xml).toContain('pid="2"');
        expect(xml).toContain('pid="3"');
        expect(xml).toContain('name="Custom1"');
        expect(xml).toContain('name="Custom2"');
        expect(xml).toContain('<vt:lpwstr>Value1</vt:lpwstr>');
        expect(xml).toContain('<vt:lpwstr>Value2</vt:lpwstr>');
      }
    });

    test('should include custom properties relationship in _rels/.rels', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          customProperties: {
            'Custom': 'Value',
          },
        },
      });

      // Verify _rels/.rels includes custom properties relationship
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const relsEntry = zipFile.entries.find((e) => e.fileName === '_rels/.rels');
      expect(relsEntry).toBeDefined();

      if (relsEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(relsEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        expect(xml).toContain('custom.xml');
        expect(xml).toMatch(/<Relationship Id="rId\d+" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/custom-properties" Target="docProps\/custom\.xml"/);
      }
    });

    test('should not write custom.xml when customProperties are empty', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          title: 'Test',
          customProperties: {},
        },
      });

      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).not.toContain('docProps/custom.xml');
    });

    test('should work with core properties and custom properties together', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
        properties: {
          title: 'Test Workbook',
          creator: 'Test',
          customProperties: {
            'Custom': 'Value',
          },
        },
      });

      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).toContain('docProps/core.xml');
      expect(fileNames).toContain('docProps/custom.xml');
    });
  });
});
