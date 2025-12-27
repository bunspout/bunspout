/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { readXlsx } from './reader';
import { writeXlsx } from './writer';

describe('XLSXReader', () => {
  const testFile = 'test-read.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should read single sheet workbook', async () => {
    // Write a test file first
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

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(1);

    const sheet = workbook.sheet('Data');
    expect(sheet.name).toBe('Data');

    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(2);
    expect(rows[0]?.cells[0]?.value).toBe('Name');
    expect(rows[1]?.cells[0]?.value).toBe('Alice');
  });

  test('should read multiple sheets workbook', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'Sheet1', rows: (async function* () { yield row([cell('A')]); })() },
        { name: 'Sheet2', rows: (async function* () { yield row([cell('B')]); })() },
      ],
    });

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(2);

    const sheet1 = workbook.sheet('Sheet1');
    expect(sheet1.name).toBe('Sheet1');

    const sheet2 = workbook.sheet(1);
    expect(sheet2.name).toBe('Sheet2');
  });

  test('should get sheet by name', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'MySheet', rows: (async function* () { yield row([cell('Test')]); })() },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('MySheet');
    expect(sheet.name).toBe('MySheet');
  });

  test('should get sheet by index', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'First', rows: (async function* () { yield row([cell('A')]); })() },
        { name: 'Second', rows: (async function* () { yield row([cell('B')]); })() },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheet0 = workbook.sheet(0);
    expect(sheet0.name).toBe('First');

    const sheet1 = workbook.sheet(1);
    expect(sheet1.name).toBe('Second');
  });

  test('should iterate all sheets', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'A', rows: (async function* () {})() },
        { name: 'B', rows: (async function* () {})() },
        { name: 'C', rows: (async function* () {})() },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheets = workbook.sheets();
    expect(sheets).toHaveLength(3);
    expect(sheets[0]?.name).toBe('A');
    expect(sheets[1]?.name).toBe('B');
    expect(sheets[2]?.name).toBe('C');
  });

  test('should handle error for invalid files', async () => {
    await expect(readXlsx('nonexistent.xlsx')).rejects.toThrow();
  });

  test('should handle error for missing sheets', async () => {
    await writeXlsx(testFile, {
      sheets: [{ name: 'Data', rows: (async function* () {})() }],
    });

    const workbook = await readXlsx(testFile);
    expect(() => workbook.sheet('NonExistent')).toThrow();
  });

  test('should read multiple sheets and verify isolation', async () => {
    // Write workbook with multiple sheets, each with different data
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Sheet1',
          rows: (async function* () {
            yield row([cell('Sheet1-Row1'), cell(1)]);
            yield row([cell('Sheet1-Row2'), cell(2)]);
          })(),
        },
        {
          name: 'Sheet2',
          rows: (async function* () {
            yield row([cell('Sheet2-Row1'), cell(10)]);
            yield row([cell('Sheet2-Row2'), cell(20)]);
            yield row([cell('Sheet2-Row3'), cell(30)]);
          })(),
        },
        {
          name: 'Sheet3',
          rows: (async function* () {
            yield row([cell('Sheet3-Only')]);
          })(),
        },
      ],
    });

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(3);

    // Read Sheet1 and verify it only contains Sheet1 data
    const sheet1 = workbook.sheet('Sheet1');
    const sheet1Rows: any[] = [];
    for await (const row of sheet1.rows()) {
      sheet1Rows.push(row);
    }
    expect(sheet1Rows).toHaveLength(2);
    expect(sheet1Rows[0]?.cells[0]?.value).toBe('Sheet1-Row1');
    expect(sheet1Rows[0]?.cells[1]?.value).toBe(1);
    expect(sheet1Rows[1]?.cells[0]?.value).toBe('Sheet1-Row2');
    expect(sheet1Rows[1]?.cells[1]?.value).toBe(2);
    // Verify no Sheet2 or Sheet3 data
    expect(sheet1Rows.every((r) => r.cells[0]?.value?.toString().startsWith('Sheet1'))).toBe(true);

    // Read Sheet2 and verify it only contains Sheet2 data
    const sheet2 = workbook.sheet('Sheet2');
    const sheet2Rows: any[] = [];
    for await (const row of sheet2.rows()) {
      sheet2Rows.push(row);
    }
    expect(sheet2Rows).toHaveLength(3);
    expect(sheet2Rows[0]?.cells[0]?.value).toBe('Sheet2-Row1');
    expect(sheet2Rows[0]?.cells[1]?.value).toBe(10);
    expect(sheet2Rows[1]?.cells[0]?.value).toBe('Sheet2-Row2');
    expect(sheet2Rows[1]?.cells[1]?.value).toBe(20);
    expect(sheet2Rows[2]?.cells[0]?.value).toBe('Sheet2-Row3');
    expect(sheet2Rows[2]?.cells[1]?.value).toBe(30);
    // Verify no Sheet1 or Sheet3 data
    expect(sheet2Rows.every((r) => r.cells[0]?.value?.toString().startsWith('Sheet2'))).toBe(true);

    // Read Sheet3 and verify it only contains Sheet3 data
    const sheet3 = workbook.sheet('Sheet3');
    const sheet3Rows: any[] = [];
    for await (const row of sheet3.rows()) {
      sheet3Rows.push(row);
    }
    expect(sheet3Rows).toHaveLength(1);
    expect(sheet3Rows[0]?.cells[0]?.value).toBe('Sheet3-Only');
    // Verify no Sheet1 or Sheet2 data
    expect(sheet3Rows.every((r) => r.cells[0]?.value?.toString().startsWith('Sheet3'))).toBe(true);
  });

  test('should read workbook with shared strings', async () => {
    // Write a file with shared strings
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

    // Read it back
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    // Verify strings are correctly resolved from shared strings table
    expect(rows).toHaveLength(2);
    expect(rows[0]?.cells[0]?.value).toBe('Hello');
    expect(rows[0]?.cells[1]?.value).toBe('World');
    expect(rows[1]?.cells[0]?.value).toBe('Hello'); // Should resolve from shared strings
    expect(rows[1]?.cells[1]?.value).toBe('Test');
  });
});

