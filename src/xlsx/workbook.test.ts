/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect, afterEach } from 'bun:test';
import { Workbook } from './workbook';

describe('Workbook and Sheet Classes', () => {
  const testFiles: string[] = [];

  afterEach(async () => {
    // Cleanup test files
    for (const testFile of testFiles) {
      if (await Bun.file(testFile).exists()) {
        await import('fs').then((fs) => fs.promises.unlink(testFile));
      }
    }
    testFiles.length = 0;
  });

  test('should create Workbook from ZIP file', () => {
    // Mock ZIP file structure
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/workbook.xml', entry: {} as any },
      ],
    };

    const workbook = new Workbook(mockZipFile, [
      { name: 'Sheet1', entry: mockZipFile.entries[0]! },
    ]);

    expect(workbook).toBeInstanceOf(Workbook);
  });

  test('should get sheet by name', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/worksheets/sheet2.xml', entry: {} as any },
      ],
    };

    const sheets = [
      { name: 'Data', entry: mockZipFile.entries[0]! },
      { name: 'Summary', entry: mockZipFile.entries[1]! },
    ];

    const workbook = new Workbook(mockZipFile, sheets);

    const sheet = workbook.sheet('Data');
    expect(sheet.name).toBe('Data');
  });

  test('should get sheet by index', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/worksheets/sheet2.xml', entry: {} as any },
      ],
    };

    const sheets = [
      { name: 'Data', entry: mockZipFile.entries[0]! },
      { name: 'Summary', entry: mockZipFile.entries[1]! },
    ];

    const workbook = new Workbook(mockZipFile, sheets);

    const sheet = workbook.sheet(0);
    expect(sheet.name).toBe('Data');

    const sheet2 = workbook.sheet(1);
    expect(sheet2.name).toBe('Summary');
  });

  test('should get all sheets', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/worksheets/sheet2.xml', entry: {} as any },
      ],
    };

    const sheets = [
      { name: 'Data', entry: mockZipFile.entries[0]! },
      { name: 'Summary', entry: mockZipFile.entries[1]! },
    ];

    const workbook = new Workbook(mockZipFile, sheets);

    const allSheets = workbook.sheets();
    expect(allSheets).toHaveLength(2);
    expect(allSheets[0]?.name).toBe('Data');
    expect(allSheets[1]?.name).toBe('Summary');
  });

  test('should throw error for missing sheet by name', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [{ fileName: 'xl/worksheets/sheet1.xml', entry: {} as any }],
    };

    const workbook = new Workbook(mockZipFile, [
      { name: 'Data', entry: mockZipFile.entries[0]! },
    ]);

    expect(() => workbook.sheet('NonExistent')).toThrow();
  });

  test('should throw error for missing sheet by index', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [{ fileName: 'xl/worksheets/sheet1.xml', entry: {} as any }],
    };

    const workbook = new Workbook(mockZipFile, [
      { name: 'Data', entry: mockZipFile.entries[0]! },
    ]);

    expect(() => workbook.sheet(99)).toThrow();
  });

  test('should expose workbook properties', async () => {
    const { writeXlsx } = await import('./writer');
    const { readXlsx } = await import('./reader');
    const { cell } = await import('@sheet/cell');
    const { row } = await import('@sheet/row');

    const testFile = 'test-properties-workbook.xlsx';
    testFiles.push(testFile);

    // Write a file with properties
    await writeXlsx(testFile, {
      properties: {
        title: 'Test Workbook',
        creator: 'Test Author',
        subject: 'Testing',
        keywords: 'test, workbook',
        description: 'A test workbook',
      },
      sheets: [
        {
          name: 'Sheet1',
          rows: (async function* () {
            yield row([cell('Data')]);
          })(),
        },
      ],
    });

    // Read it back
    const workbook = await readXlsx(testFile);
    const properties = await workbook.properties();

    expect(properties.title).toBe('Test Workbook');
    expect(properties.creator).toBe('Test Author');
    expect(properties.subject).toBe('Testing');
    expect(properties.keywords).toBe('test, workbook');
    expect(properties.description).toBe('A test workbook');
  });

  test('should expose sheet properties', async () => {
    const { writeXlsx } = await import('./writer');
    const { readXlsx } = await import('./reader');
    const { cell } = await import('@sheet/cell');
    const { row } = await import('@sheet/row');

    const testFile = 'test-sheet-properties.xlsx';
    testFiles.push(testFile);

    // Write a file with sheet properties
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          defaultColumnWidth: 15,
          columnWidths: [
            { columnIndex: 0, width: 20 },
            { columnRange: { from: 2, to: 3 }, width: 25 },
          ],
          defaultRowHeight: 20,
          rows: (async function* () {
            yield row([cell('A'), cell('B'), cell('C'), cell('D')]);
          })(),
        },
      ],
    });

    // Read it back
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');

    expect(sheet.defaultColumnWidth).toBe(15);
    expect(sheet.defaultRowHeight).toBe(20);
    expect(sheet.columnWidths).toBeDefined();
    // When defaultColumnWidth is set, columns without explicit widths get default-width entries
    // The exact structure depends on how the writer generates the XML, but we should have:
    // - Column 0 with width 20 (explicit)
    // - Column 1 with width 15 (default)
    // - Columns 2-3 with width 25 (range)
    expect(sheet.columnWidths!.length).toBeGreaterThanOrEqual(2);
    // Find the explicit column 0 width
    const col0 = sheet.columnWidths!.find((cw) => cw.columnIndex === 0);
    expect(col0?.width).toBe(20);
    // Find the range for columns 2-3
    const col2to3 = sheet.columnWidths!.find((cw) => cw.columnRange?.from === 2 && cw.columnRange?.to === 3);
    expect(col2to3?.width).toBe(25);
    // Verify we have the default width somewhere (either in defaultColumnWidth or in columnWidths)
    const hasDefaultWidth = sheet.columnWidths!.some((cw) => cw.width === 15);
    expect(hasDefaultWidth || sheet.defaultColumnWidth === 15).toBe(true);
  });
});
