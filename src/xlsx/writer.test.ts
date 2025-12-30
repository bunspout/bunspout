/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { readXlsx } from './reader';
import { writeXlsx } from './writer';

describe('XLSXWriter', () => {
  const testFile = 'test-output.xlsx';

  afterEach(async () => {
    // Clean up test file after each test
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should write single sheet workbook', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            yield row([cell('Name'), cell('Age')]);
            yield row([cell('Alice'), cell(30)]);
          })(),
        },
      ],
    });

    expect(await Bun.file(testFile).exists()).toBe(true);
  });

  test('should write multiple sheets workbook', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Sheet1',
          rows: (async function* () {
            yield row([cell('A')]);
          })(),
        },
        {
          name: 'Sheet2',
          rows: (async function* () {
            yield row([cell('B')]);
          })(),
        },
      ],
    });

    expect(await Bun.file(testFile).exists()).toBe(true);

    // Verify we can read it back
    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(2);
    expect(workbook.sheet('Sheet1').name).toBe('Sheet1');
    expect(workbook.sheet('Sheet2').name).toBe('Sheet2');

    // Verify sheet contents
    const sheet1 = workbook.sheet('Sheet1');
    const sheet1Rows: any[] = [];
    for await (const row of sheet1.rows()) {
      sheet1Rows.push(row);
    }
    expect(sheet1Rows).toHaveLength(1);
    expect(sheet1Rows[0]?.cells[0]?.value).toBe('A');

    const sheet2 = workbook.sheet('Sheet2');
    const sheet2Rows: any[] = [];
    for await (const row of sheet2.rows()) {
      sheet2Rows.push(row);
    }
    expect(sheet2Rows).toHaveLength(1);
    expect(sheet2Rows[0]?.cells[0]?.value).toBe('B');

  });

  test('should write empty sheet', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Empty',
          rows: (async function* () {
            // No rows
          })(),
        },
      ],
    });

    expect(await Bun.file(testFile).exists()).toBe(true);

    // Verify sheet exists and can be read
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Empty');
    expect(sheet).toBeDefined();
    expect(sheet.name).toBe('Empty');

  });

  test('should verify sheet names are correct', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'My Custom Sheet',
          rows: (async function* () {
            yield row([cell('Test')]);
          })(),
        },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('My Custom Sheet');
    expect(sheet.name).toBe('My Custom Sheet');

  });

  test('should handle streaming behavior with large dataset', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Large',
          rows: (async function* () {
            for (let i = 0; i < 100; i++) {
              yield row([cell(`Row${i}`), cell(i)]);
            }
          })(),
        },
      ],
    });

    expect(await Bun.file(testFile).exists()).toBe(true);

    // Verify we can read it back
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Large');
    let count = 0;
    for await (const row of sheet.rows()) {
      count++;
      expect(row.cells[0]?.value).toBe(`Row${count - 1}`);
    }
    expect(count).toBe(100);

  });

  test('should support WriterOptions.sharedStrings inline (default)', async () => {
    await writeXlsx(
      testFile,
      {
        sheets: [
          {
            name: 'Test',
            rows: (async function* () {
              yield row([cell('Hello')]);
            })(),
          },
        ],
      },
      { sharedStrings: 'inline' },
    );

    expect(await Bun.file(testFile).exists()).toBe(true);

  });

  test('should preserve newlines and handle control characters in round-trip', async () => {
    const testData = [
      row([cell('Line1\nLine2\r\nLine3')]), // Newlines
      row([cell('Tab\tSeparated')]), // Tabs
      row([cell('Normal text')]), // Normal text
    ];

    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            for (const row of testData) yield row;
          })(),
        },
      ],
    });

    // Read back and verify data is preserved
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('Line1\nLine2\r\nLine3');
    expect(rows[1]?.cells[0]?.value).toBe('Tab\tSeparated');
    expect(rows[2]?.cells[0]?.value).toBe('Normal text');

  });

  test('should remove invalid control characters while preserving valid ones', async () => {
    // Create string with invalid control chars (null, bell) and valid ones (tab, newline, CR)
    const textWithControlChars =
      'Start' +
      String.fromCharCode(0x00) + // null - invalid
      'Middle' +
      String.fromCharCode(0x07) + // bell - invalid
      String.fromCharCode(0x09) + // tab - valid
      'AfterTab' +
      String.fromCharCode(0x0A) + // newline - valid
      'AfterNewline' +
      String.fromCharCode(0x0D) + // carriage return - valid
      'End';

    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            yield row([cell(textWithControlChars)]);
          })(),
        },
      ],
    });

    // Read back - invalid chars should be removed, valid ones preserved
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const result = rows[0]?.cells[0]?.value as string;

    // Should contain the text parts
    expect(result).toContain('Start');
    expect(result).toContain('Middle');
    expect(result).toContain('AfterTab');
    expect(result).toContain('AfterNewline');
    expect(result).toContain('End');

    // Should preserve valid control chars
    expect(result).toContain('\t'); // tab
    expect(result).toContain('\n'); // newline
    expect(result).toContain('\r'); // carriage return

    // Should NOT contain invalid control chars
    expect(result).not.toContain(String.fromCharCode(0x00));
    expect(result).not.toContain(String.fromCharCode(0x07));

  });

  test('should write rows with empty cells', async () => {
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            yield row([cell('A'), cell(''), cell('C')]); // Empty cell in middle
            yield row([cell(''), cell('B'), cell('')]); // Empty cells at start and end
            yield row([cell(''), cell(''), cell('')]); // All empty cells
          })(),
        },
      ],
    });

    expect(await Bun.file(testFile).exists()).toBe(true);

    // Read back and verify empty cells are preserved
    const workbook = await readXlsx(testFile, { skipEmptyRows: false });
    const sheet = workbook.sheet('Data');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(3);
    expect(rows[0]?.cells).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('A');
    expect(rows[0]?.cells[1]?.value).toBe(''); // Empty cell
    expect(rows[0]?.cells[2]?.value).toBe('C');

    expect(rows[1]?.cells[0]?.value).toBe(''); // Empty at start
    expect(rows[1]?.cells[1]?.value).toBe('B');
    expect(rows[1]?.cells[2]?.value).toBe(''); // Empty at end

    expect(rows[2]?.cells).toHaveLength(3);
    expect(rows[2]?.cells[0]?.value).toBe(''); // All empty
    expect(rows[2]?.cells[1]?.value).toBe('');
    expect(rows[2]?.cells[2]?.value).toBe('');

  });

  describe('File Validation', () => {
    test('should generate valid XML for empty sheets', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Empty',
            rows: (async function* () {
              // No rows
            })(),
          },
        ],
      });

      expect(await Bun.file(testFile).exists()).toBe(true);

      // Read back and verify the sheet XML is valid
      const workbook = await readXlsx(testFile);
      const sheet = workbook.sheet('Empty');
      const rows: any[] = [];
      for await (const row of sheet.rows()) {
        rows.push(row);
      }

      // Empty sheet should have no rows but be readable
      expect(rows).toHaveLength(0);
      expect(sheet.name).toBe('Empty');

    });

    test('should generate valid ZIP file structure', async () => {
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

      // Verify file exists and is not empty
      const file = Bun.file(testFile);
      expect(await file.exists()).toBe(true);
      const size = (await file.arrayBuffer()).byteLength;
      expect(size).toBeGreaterThan(0);

      // Verify ZIP magic bytes (PK header)
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipMagic = buffer.slice(0, 2).toString('ascii');
      expect(zipMagic).toBe('PK'); // ZIP files start with "PK"

    });

    test('should contain all required XLSX files in ZIP', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Sheet1',
            rows: (async function* () {
              yield row([cell('A')]);
            })(),
          },
        ],
      });

      // Read the ZIP and verify required files are present
      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);

      const fileNames = zipFile.entries.map((e) => e.fileName);

      // Required files for XLSX
      expect(fileNames).toContain('[Content_Types].xml');
      expect(fileNames).toContain('_rels/.rels');
      expect(fileNames).toContain('xl/workbook.xml');
      expect(fileNames).toContain('xl/_rels/workbook.xml.rels');
      expect(fileNames).toContain('xl/worksheets/sheet1.xml');

    });

    test('should generate well-formed XML files', async () => {
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

      // Read back - if XML is malformed, this will throw
      const workbook = await readXlsx(testFile);
      expect(workbook).toBeDefined();
      expect(workbook.sheets()).toHaveLength(1);

      // Try to read the sheet - if XML is malformed, parsing will fail
      const sheet = workbook.sheet('Data');
      const rows: any[] = [];
      for await (const row of sheet.rows()) {
        rows.push(row);
      }
      expect(rows).toHaveLength(1);

    });

    test('should generate file with correct MIME type', async () => {
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

      const file = Bun.file(testFile);

      // XLSX files are ZIP files, so they should have application/zip or
      // application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
      // Bun may detect it as application/zip or application/x-zip-compressed
      const type = file.type;
      expect(type).toMatch(/zip|spreadsheetml|octet-stream/);

      // Verify it's a valid ZIP by checking magic bytes
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipMagic = buffer.slice(0, 2).toString('ascii');
      expect(zipMagic).toBe('PK');

    });

    test('should generate valid XLSX for multiple sheets', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Sheet1',
            rows: (async function* () {
              yield row([cell('A')]);
            })(),
          },
          {
            name: 'Sheet2',
            rows: (async function* () {
              yield row([cell('B')]);
            })(),
          },
          {
            name: 'Empty',
            rows: (async function* () {})(),
          },
        ],
      });

      // Verify file can be read back
      const workbook = await readXlsx(testFile);
      expect(workbook.sheets()).toHaveLength(3);

      // Verify all sheets are accessible
      expect(workbook.sheet('Sheet1').name).toBe('Sheet1');
      expect(workbook.sheet('Sheet2').name).toBe('Sheet2');
      expect(workbook.sheet('Empty').name).toBe('Empty');

      // Verify ZIP contains all required files
      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).toContain('xl/worksheets/sheet1.xml');
      expect(fileNames).toContain('xl/worksheets/sheet2.xml');
      expect(fileNames).toContain('xl/worksheets/sheet3.xml');

    });

    test('should generate valid XML with special characters in sheet names', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Sheet & Data <Test>',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
      });

      // Verify file can be read back and sheet name is correct
      const workbook = await readXlsx(testFile);
      const sheet = workbook.sheet('Sheet & Data <Test>');
      expect(sheet.name).toBe('Sheet & Data <Test>');

      // Verify XML is well-formed (if it wasn't, reading would fail)
      const rows: any[] = [];
      for await (const row of sheet.rows()) {
        rows.push(row);
      }
      expect(rows).toHaveLength(1);

    });
  });

  describe('Shared Strings', () => {
    test('should write using shared strings when option is enabled', async () => {
      await writeXlsx(
        testFile,
        {
          sheets: [
            {
              name: 'Data',
              rows: (async function* () {
                yield row([cell('Hello'), cell('World')]);
                yield row([cell('Hello'), cell('Test')]); // Duplicate string
              })(),
            },
          ],
        },
        { sharedStrings: 'shared' },
      );

      const file = Bun.file(testFile);
      expect(await file.exists()).toBe(true);

      // Verify shared strings file exists in ZIP
      const { openZip } = await import('../zip/reader');
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).toContain('xl/sharedStrings.xml');

      // Verify we can read it back
      const workbook = await readXlsx(testFile);
      const sheet = workbook.sheet('Data');
      const rows: any[] = [];
      for await (const row of sheet.rows()) {
        rows.push(row);
      }

      expect(rows).toHaveLength(2);
      expect(rows[0]?.cells[0]?.value).toBe('Hello');
      expect(rows[0]?.cells[1]?.value).toBe('World');
      expect(rows[1]?.cells[0]?.value).toBe('Hello');
      expect(rows[1]?.cells[1]?.value).toBe('Test');

    });

    test('should deduplicate strings in shared strings table', async () => {
      await writeXlsx(
        testFile,
        {
          sheets: [
            {
              name: 'Data',
              rows: (async function* () {
                yield row([cell('Duplicate')]);
                yield row([cell('Duplicate')]);
                yield row([cell('Duplicate')]);
                yield row([cell('Other')]);
              })(),
            },
          ],
        },
        { sharedStrings: 'shared' },
      );

      // Read shared strings XML to verify deduplication
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const sharedStringsEntry = zipFile.entries.find(
        (e) => e.fileName === 'xl/sharedStrings.xml',
      );
      expect(sharedStringsEntry).toBeDefined();

      if (sharedStringsEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(sharedStringsEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Should have uniqueCount="2" (Duplicate and Other)
        expect(xml).toContain('uniqueCount="2"');
        // Should have count="4" (total string references)
        expect(xml).toContain('count="4"');
        expect(xml).toContain('<si><t>Duplicate</t></si>');
        expect(xml).toContain('<si><t>Other</t></si>');
      }

    });

    test('should work with multiple sheets using shared strings', async () => {
      await writeXlsx(
        testFile,
        {
          sheets: [
            {
              name: 'Sheet1',
              rows: (async function* () {
                yield row([cell('Shared')]);
                yield row([cell('Sheet1-Only')]);
              })(),
            },
            {
              name: 'Sheet2',
              rows: (async function* () {
                yield row([cell('Shared')]); // Same string as Sheet1
                yield row([cell('Sheet2-Only')]);
              })(),
            },
          ],
        },
        { sharedStrings: 'shared' },
      );

      // Verify shared strings contains all unique strings from both sheets
      const { openZip, readZipEntry } = await import('../zip/reader');
      const { bytesToString } = await import('../adapters/common');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const sharedStringsEntry = zipFile.entries.find(
        (e) => e.fileName === 'xl/sharedStrings.xml',
      );

      if (sharedStringsEntry) {
        const xmlBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(sharedStringsEntry, zipFile.zipFile)) {
          xmlBytes.push(chunk);
        }
        const xml = await bytesToString(async function* () {
          for (const chunk of xmlBytes) yield chunk;
        }());

        // Should have 3 unique strings: Shared, Sheet1-Only, Sheet2-Only
        expect(xml).toContain('uniqueCount="3"');
        expect(xml).toContain('<si><t>Shared</t></si>');
        expect(xml).toContain('<si><t>Sheet1-Only</t></si>');
        expect(xml).toContain('<si><t>Sheet2-Only</t></si>');
      }

      // Verify we can read both sheets back
      const workbook = await readXlsx(testFile);
      expect(workbook.sheets()).toHaveLength(2);

      const sheet1 = workbook.sheet('Sheet1');
      const rows1: any[] = [];
      for await (const row of sheet1.rows()) {
        rows1.push(row);
      }
      expect(rows1[0]?.cells[0]?.value).toBe('Shared');
      expect(rows1[1]?.cells[0]?.value).toBe('Sheet1-Only');

      const sheet2 = workbook.sheet('Sheet2');
      const rows2: any[] = [];
      for await (const row of sheet2.rows()) {
        rows2.push(row);
      }
      expect(rows2[0]?.cells[0]?.value).toBe('Shared');
      expect(rows2[1]?.cells[0]?.value).toBe('Sheet2-Only');

    });

    test('should default to inline strings when not specified', async () => {
      await writeXlsx(testFile, {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Hello')]);
            })(),
          },
        ],
      });

      // Verify shared strings file does NOT exist
      const { openZip } = await import('../zip/reader');
      const file = Bun.file(testFile);
      const buffer = Buffer.from(await file.arrayBuffer());
      const zipFile = await openZip(buffer);
      const fileNames = zipFile.entries.map((e) => e.fileName);

      expect(fileNames).not.toContain('xl/sharedStrings.xml');

    });
  });
});
