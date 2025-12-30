/* eslint-disable @typescript-eslint/no-explicit-any */
import { strictEqual, deepStrictEqual } from 'node:assert';
import { access, unlink } from 'node:fs/promises';
import { describe, test, afterEach } from 'node:test';
import { readXlsx } from '@xlsx/reader';
import { writeXlsx } from '@xlsx/writer';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';

// Helper to check if file exists (Node.js equivalent of Bun.file().exists())
async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

// Helper to match Bun's expect API for easier migration
const expect = {
  toHaveLength: (actual: any, expected: number) => {
    strictEqual(actual.length, expected, `Expected length ${expected}, got ${actual.length}`);
  },
  toBe: (actual: any, expected: any) => {
    strictEqual(actual, expected);
  },
  toEqual: (actual: any, expected: any) => {
    deepStrictEqual(actual, expected);
  },
  toBeInstanceOf: (actual: any, expected: any) => {
    if (!(actual instanceof expected)) {
      throw new Error(`Expected instance of ${expected.name}, got ${typeof actual}`);
    }
  },
};

describe('Integration Tests (Node.js)', () => {
  const testFile = 'integration-test-node.xlsx';

  afterEach(async () => {
    if (await fileExists(testFile)) {
      await unlink(testFile);
    }
    // Also clean up date-test.xlsx if it exists
    if (await fileExists('date-test-node.xlsx')) {
      await unlink('date-test-node.xlsx');
    }
  });

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
    expect.toHaveLength(readRows, originalRows.length);
    for (let i = 0; i < originalRows.length; i++) {
      expect.toHaveLength(readRows[i]?.cells, originalRows[i]!.cells.length);
      for (let j = 0; j < originalRows[i]!.cells.length; j++) {
        expect.toBe(
          readRows[i]?.cells[j]?.value,
          originalRows[i]!.cells[j]!.value,
        );
        expect.toBe(
          readRows[i]?.cells[j]?.type,
          originalRows[i]!.cells[j]?.type,
        );
      }
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

    const workbook = await readXlsx(testFile, { skipEmptyRows: false });
    const sheet = workbook.sheet('Test');

    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readRows.push(row);
    }

    expect.toHaveLength(readRows, rows.length);
    expect.toHaveLength(readRows[0]?.cells, 0);
    expect.toBe(readRows[1]?.cells[0]?.value, 'String');
    expect.toBe(readRows[2]?.cells[0]?.value, 42);
    expect.toHaveLength(readRows[3]?.cells, 3);
    expect.toBe(readRows[3]?.rowIndex, 4);
  });

  test('should parse date cells with 1900 date system (use1904Dates: false)', async () => {
    const dateTestFile = 'date-test-node.xlsx';

    // Write a file with date cells (this will be written as numbers with date type)
    await writeXlsx(dateTestFile, {
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

    // Read back using the 1900 date system (use1904Dates: false)
    const workbook = await readXlsx(dateTestFile, { use1904Dates: false });
    const sheet = workbook.sheet('Dates');

    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readRows.push(row);
    }

    expect.toHaveLength(readRows, 2);

    // Check that the date cell has the parsed Date directly in value
    const dateCell = readRows[1]?.cells[0];
    expect.toBe(dateCell?.type, 'date');
    expect.toBeInstanceOf(dateCell?.value, Date);
    expect.toBe((dateCell?.value as Date)?.getFullYear(), 2024);
    expect.toBe((dateCell?.value as Date)?.getMonth(), 0); // January
    expect.toBe((dateCell?.value as Date)?.getDate(), 15);
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

    expect.toBe(writeCount, 100);

    // Stream read
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Large');
    let readCount = 0;
    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readCount++;
      readRows.push(row);
    }

    expect.toBe(readCount, 100);
    expect.toHaveLength(readRows, 100);
    expect.toBe(readRows[0]?.cells[0]?.value, 'Row0');
    expect.toBe(readRows[99]?.cells[0]?.value, 'Row99');
  });

  test('should maintain consistency across multiple sheet iterations', async () => {
    // Create a test file with multiple sheets containing diverse data types
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Sheet1',
          rows: (async function* () {
            yield row([cell('Name'), cell('Age'), cell('Salary')]);
            yield row([cell('Alice'), cell(30), cell(50000.50)]);
            yield row([cell('Bob'), cell(25), cell(45000)]);
            yield row([cell(''), cell(null), cell('')]); // Empty cells
          })(),
        },
        {
          name: 'Sheet2',
          rows: (async function* () {
            yield row([cell('Product'), cell('Price'), cell('Date')]);
            yield row([cell('Laptop'), cell(1200.99), cell(new Date('2024-01-15'))]);
            yield row([cell('Mouse'), cell(25.50), cell(new Date('2024-02-20'))]);
            yield row([cell(''), cell(0), cell(null)]); // Mixed empty cells
            yield row([cell('Keyboard'), cell(75), cell(new Date('2024-03-10'))]);
          })(),
        },
        {
          name: 'Sheet3',
          rows: (async function* () {
            yield row([cell('ID'), cell('Status')]);
            yield row([cell(1), cell('Active')]);
            yield row([cell(2), cell('Inactive')]);
            // Empty row follows
            yield row([], { rowIndex: 4 }); // Explicitly empty row
            yield row([cell(3), cell('Pending')]);
          })(),
        },
      ],
    });

    // First iteration: collect all sheet data
    const workbook1 = await readXlsx(testFile, { skipEmptyRows: false });
    const firstIterationData: Array<{
      name: string;
      index: number;
      rows: Array<{
        rowIndex: number;
        cells: Array<{ value: any; type: string }>;
      }>;
    }> = [];

    let sheetIndex = 0;
    for (const sheet of workbook1.sheets()) {
      const sheetData = {
        name: sheet.name,
        index: sheetIndex,
        rows: [] as Array<{
          rowIndex: number;
          cells: Array<{ value: any; type: string }>;
        }>,
      };

      for await (const row of sheet.rows()) {
        sheetData.rows.push({
          rowIndex: row.rowIndex ?? 0,
          cells: row.cells.map(cell => ({
            value: cell?.value,
            type: cell?.type ?? 'unknown',
          })),
        });
      }

      firstIterationData.push(sheetData);
      sheetIndex++;
    }

    // Second iteration: collect all sheet data again
    const workbook2 = await readXlsx(testFile, { skipEmptyRows: false });
    const secondIterationData: Array<{
      name: string;
      index: number;
      rows: Array<{
        rowIndex: number;
        cells: Array<{ value: any; type: string }>;
      }>;
    }> = [];

    sheetIndex = 0;
    for (const sheet of workbook2.sheets()) {
      const sheetData = {
        name: sheet.name,
        index: sheetIndex,
        rows: [] as Array<{
          rowIndex: number;
          cells: Array<{ value: any; type: string }>;
        }>,
      };

      for await (const row of sheet.rows()) {
        sheetData.rows.push({
          rowIndex: row.rowIndex ?? 0,
          cells: row.cells.map(cell => ({
            value: cell?.value,
            type: cell?.type ?? 'unknown',
          })),
        });
      }

      secondIterationData.push(sheetData);
      sheetIndex++;
    }

    // Assertions: Verify complete consistency between iterations

    // Same number of sheets
    expect.toHaveLength(secondIterationData, firstIterationData.length);

    // Each sheet has same name and index
    for (let i = 0; i < firstIterationData.length; i++) {
      const firstSheet = firstIterationData[i]!;
      const secondSheet = secondIterationData[i]!;

      expect.toBe(secondSheet.name, firstSheet.name);
      expect.toBe(secondSheet.index, firstSheet.index);

      // Same number of rows
      expect.toHaveLength(secondSheet.rows, firstSheet.rows.length);

      // Each row has same data
      for (let j = 0; j < firstSheet.rows.length; j++) {
        const firstRow = firstSheet.rows[j]!;
        const secondRow = secondSheet.rows[j]!;

        expect.toBe(secondRow.rowIndex, firstRow.rowIndex);
        expect.toHaveLength(secondRow.cells, firstRow.cells.length);

        // Each cell has same value and type
        for (let k = 0; k < firstRow.cells.length; k++) {
          const firstCell = firstRow.cells[k]!;
          const secondCell = secondRow.cells[k]!;

          expect.toEqual(secondCell.value, firstCell.value);
          expect.toBe(secondCell.type, firstCell.type);
        }
      }
    }
  });
});
