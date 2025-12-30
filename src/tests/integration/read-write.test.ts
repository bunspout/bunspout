/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect, afterEach } from 'bun:test';
import { readXlsx } from '@xlsx/reader';
import { writeXlsx } from '@xlsx/writer';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';

describe('Integration Tests', () => {
  const testFile = 'integration-test.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
    // Also clean up date-test.xlsx if it exists
    if (await Bun.file('date-test.xlsx').exists()) {
      await import('fs').then((fs) => fs.promises.unlink('date-test.xlsx'));
    }
    // Clean up date-test-1904.xlsx if it exists
    if (await Bun.file('date-test-1904.xlsx').exists()) {
      await import('fs').then((fs) => fs.promises.unlink('date-test-1904.xlsx'));
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

    expect(readRows).toHaveLength(rows.length);
    expect(readRows[0]?.cells).toHaveLength(0);
    expect(readRows[1]?.cells[0]?.value).toBe('String');
    expect(readRows[2]?.cells[0]?.value).toBe(42);
    expect(readRows[3]?.cells).toHaveLength(3);
    expect(readRows[3]?.rowIndex).toBe(4);
  });

  test('should parse date cells with 1900 date system (use1904Dates: false)', async () => {
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

    // Read back using the 1900 date system (use1904Dates: false)
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
  });

  test('should parse date cells with 1904 date system (use1904Dates: true)', async () => {
    const testFile = 'date-test-1904.xlsx';

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

    // Read back using the 1904 date system (use1904Dates: true)
    const workbook = await readXlsx(testFile, { use1904Dates: true });
    const sheet = workbook.sheet('Dates');

    const readRows: any[] = [];
    for await (const row of sheet.rows()) {
      readRows.push(row);
    }

    expect(readRows).toHaveLength(2);

    // Check that the date cell has the parsed Date directly in value
    // Note: The date will be interpreted differently with 1904 system
    const dateCell = readRows[1]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBeInstanceOf(Date);
    // With 1904 system, the same Excel serial number represents a different date
    // We verify it's a valid date object
    expect((dateCell?.value as Date)?.getTime()).toBeGreaterThan(0);
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
    expect(secondIterationData).toHaveLength(firstIterationData.length);

    // Each sheet has same name and index
    for (let i = 0; i < firstIterationData.length; i++) {
      const firstSheet = firstIterationData[i]!;
      const secondSheet = secondIterationData[i]!;

      expect(secondSheet.name).toBe(firstSheet.name);
      expect(secondSheet.index).toBe(firstSheet.index);

      // Same number of rows
      expect(secondSheet.rows).toHaveLength(firstSheet.rows.length);

      // Each row has same data
      for (let j = 0; j < firstSheet.rows.length; j++) {
        const firstRow = firstSheet.rows[j]!;
        const secondRow = secondSheet.rows[j]!;

        expect(secondRow.rowIndex).toBe(firstRow.rowIndex);
        expect(secondRow.cells).toHaveLength(firstRow.cells.length);

        // Each cell has same value and type
        for (let k = 0; k < firstRow.cells.length; k++) {
          const firstCell = firstRow.cells[k]!;
          const secondCell = secondRow.cells[k]!;

          expect(secondCell.value).toEqual(firstCell.value);
          expect(secondCell.type).toBe(firstCell.type);
        }
      }
    }
  });
});
