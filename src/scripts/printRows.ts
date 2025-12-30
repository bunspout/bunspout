import { readXlsx } from '@xlsx/reader';

async function main() {
  const filePath = Bun.argv[2];

  if (!filePath) {
    console.error('Usage: bun run src/scripts/printRows.ts <path-to-xlsx-file>');
    process.exit(1);
  }

  try {
    const workbook = await readXlsx(filePath);
    const sheet = workbook.sheet(0); // Get first sheet

    console.log(`Sheet: ${sheet.name}\n`);

    let rowCount = 0;
    for await (const row of sheet.rows()) {
      rowCount++;
      const rowIndex = row.rowIndex ?? rowCount;
      const cellValues = row.cells.map((cell) => {
        if (!cell || cell.value === null || cell.value === undefined) {
          return '';
        }
        return String(cell.value);
      });
      console.log(`Row ${rowIndex}: ${cellValues.join(' | ')}`);
    }

    console.log(`\nTotal rows: ${rowCount}`);
  } catch (error) {
    console.error('Error reading XLSX file:', error instanceof Error ? error.message : error);
    process.exit(1);
  }
}

main();
