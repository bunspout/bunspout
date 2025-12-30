import { writeXlsx } from '@xlsx/writer';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';

async function main() {
  const rowCount = parseInt(Bun.argv[2] || '0', 10);
  const colCount = parseInt(Bun.argv[3] || '0', 10);
  const filePath = Bun.argv[4] || 'test-data.xlsx';

  if (!rowCount || !colCount || rowCount <= 0 || colCount <= 0) {
    console.error('Usage: bun run src/scripts/writeTestData.ts <rowCount> <colCount> [output-file]');
    console.error('Example: bun run src/scripts/writeTestData.ts 100 10 test.xlsx');
    process.exit(1);
  }

  try {
    console.log(`Generating ${rowCount} rows × ${colCount} columns...`);

    await writeXlsx(filePath, {
      sheets: [
        {
          name: 'Sheet1',
          rows: (async function* () {
            for (let rowIdx = 0; rowIdx < rowCount; rowIdx++) {
              const cells = [];
              for (let colIdx = 0; colIdx < colCount; colIdx++) {
                cells.push(cell(`xlsx-${rowIdx}-${colIdx}`));
              }
              yield row(cells, { rowIndex: rowIdx + 1 });
            }
          })(),
        },
      ],
    });

    console.log(`✓ Successfully wrote ${filePath}`);
    console.log(`  Rows: ${rowCount}, Columns: ${colCount}`);
  } catch (error) {
    console.error('Error writing XLSX file:', error instanceof Error ? error.message : error);
    process.exit(1);
  }
}

main();
