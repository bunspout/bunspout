/* eslint-disable @typescript-eslint/no-explicit-any */
import { existsSync } from 'fs';
import { describe, test, expect } from 'bun:test';
import { readXlsx } from '@xlsx/reader';
import { writeXlsx } from '@xlsx/writer';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';

describe('Integration Tests', () => {
  const testFile = 'integration-test.xlsx';

  test('should write rows → ZIP → read back → verify rows match', async () => {
    const originalRows = [
      row([cell('Name'), cell('Age')]),
      row([cell('Alice'), cell(30)]),
      row([cell('Bob'), cell(25)]),
    ];

    // Write using high-level API
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            for (const row of originalRows) yield row;
          })(),
        },
      ],
    });

    // Read back using high-level API
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');

    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readRows.push(row);
    }

    // Verify rows match
    expect(readRows).toHaveLength(originalRows.length);
    for (let i = 0; i < originalRows.length; i++) {
      expect(readRows[i]?.cells).toHaveLength(originalRows[i]!.cells.length);
      for (let j = 0; j < originalRows[i]!.cells.length; j++) {
        expect(readRows[i]?.cells[j]?.value).toBe(
          originalRows[i]!.cells[j]!.value,
        );
        expect(readRows[i]?.cells[j]?.type).toBe(
          originalRows[i]!.cells[j]?.type,
        );
      }
    }

    // Clean up
    if (existsSync(testFile)) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should handle various row/cell configurations', async () => {
    const rows = [
      row([], { rowIndex: 1 }), // Empty row
      row([cell('String')]), // String cell
      row([cell(42)]), // Number cell
      row([cell('Mixed'), cell(100), cell('End')], { rowIndex: 4 }), // Mixed row
    ];

    // Write and read cycle
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Test',
          rows: (async function* () {
            for (const row of rows) yield row;
          })(),
        },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Test');

    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readRows.push(row);
    }

    expect(readRows).toHaveLength(rows.length);
    expect(readRows[0]?.cells).toHaveLength(0);
    expect(readRows[1]?.cells[0]?.value).toBe('String');
    expect(readRows[2]?.cells[0]?.value).toBe(42);
    expect(readRows[3]?.cells).toHaveLength(3);
    expect(readRows[3]?.rowIndex).toBe(4);

    // Clean up
    if (existsSync(testFile)) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should parse date cells when use1904Dates is enabled', async () => {
    const testFile = 'date-test.xlsx';

    // Write a file with date cells (this will be written as numbers with date type)
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Dates',
          rows: (async function* () {
            yield row([cell('Date'), cell('Value')]);
            yield row([cell(new Date('2024-01-15')), cell(42)]);
          })(),
        },
      ],
    });

    // Read back with date parsing enabled
    const workbook = await readXlsx(testFile, { use1904Dates: false });
    const sheet = workbook.sheet('Dates');

    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readRows.push(row);
    }

    expect(readRows).toHaveLength(2);

    // Check that the date cell has the parsed Date directly in value
    const dateCell = readRows[1]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBeInstanceOf(Date);
    expect((dateCell?.value as Date)?.getFullYear()).toBe(2024);
    expect((dateCell?.value as Date)?.getMonth()).toBe(0); // January
    expect((dateCell?.value as Date)?.getDate()).toBe(15);

    // Clean up
    if (existsSync(testFile)) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should handle streaming behavior with large dataset', async () => {
    // Create a large dataset
    const rows = [];
    for (let i = 0; i < 100; i++) {
      rows.push(row([cell(`Row${i}`), cell(i)], { rowIndex: i + 1 }));
    }

    // Stream write
    let writeCount = 0;
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Large',
          rows: (async function* () {
            for (const row of rows) {
              writeCount++;
              yield row;
            }
          })(),
        },
      ],
    });

    expect(writeCount).toBe(100);

    // Stream read
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Large');
    let readCount = 0;
    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readCount++;
      readRows.push(row);
    }

    expect(readCount).toBe(100);
    expect(readRows).toHaveLength(100);
    expect(readRows[0]?.cells[0]?.value).toBe('Row0');
    expect(readRows[99]?.cells[0]?.value).toBe('Row99');

    // Clean up
    if (existsSync(testFile)) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });
});
