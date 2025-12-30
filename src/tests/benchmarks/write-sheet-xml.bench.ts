import assert from 'node:assert';
import { Bench } from 'tinybench';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';
import { writeSheetXml } from '@xml/writer';
import type { WriteSheetXmlOptions } from '@xml/writer';
import type { Row } from '../../types';

const ROW_COUNT = 100_000;
const COLUMN_COUNT = 10;

/**
 * Generate test rows
 */
async function* generateRows(): AsyncGenerator<Row> {
  for (let i = 1; i <= ROW_COUNT; i++) {
    const cells = [];
    for (let j = 0; j < COLUMN_COUNT; j++) {
      cells.push(cell(`Row ${i} Col ${j}`));
    }
    yield row(cells, { rowIndex: i });
  }
}

/**
 * Consume the async iterable to measure performance
 */
async function consumeXml(xmlStream: AsyncIterable<string>): Promise<number> {
  let bytes = 0;
  for await (const chunk of xmlStream) {
    bytes += chunk.length;
  }
  return bytes;
}

const bench = new Bench({ iterations: 10 });

// Variant 1: No widths (fast path - streaming)
bench.add('writeSheetXml - no widths', async () => {
  const start = performance.now();
  const rows = generateRows();
  const xmlStream = writeSheetXml(rows);
  const bytes = await consumeXml(xmlStream);
  const time = performance.now() - start;

  assert(bytes > 0, 'Should generate XML');
  return time;
});

// Variant 2: Default width only (fast path - streaming with sheetFormatPr)
bench.add('writeSheetXml - default width only', async () => {
  const start = performance.now();
  const rows = generateRows();
  const options: WriteSheetXmlOptions = {
    columnWidths: {
      defaultColumnWidth: 15,
    },
  };
  const xmlStream = writeSheetXml(rows, options);
  const bytes = await consumeXml(xmlStream);
  const time = performance.now() - start;

  assert(bytes > 0, 'Should generate XML');
  return time;
});

// Variant 3: Per-column widths (slow path - buffering)
bench.add('writeSheetXml - per-column widths', async () => {
  const start = performance.now();
  const rows = generateRows();
  const options: WriteSheetXmlOptions = {
    columnWidths: {
      columnWidths: [
        { columnIndex: 0, width: 20 },
        { columnIndex: 1, width: 15 },
        { columnIndex: 2, width: 25 },
        { columnRange: { from: 3, to: 9 }, width: 12 },
      ],
    },
  };
  const xmlStream = writeSheetXml(rows, options);
  const bytes = await consumeXml(xmlStream);
  const time = performance.now() - start;

  assert(bytes > 0, 'Should generate XML');
  return time;
});

// Variant 4: Auto-detect (slow path - buffering)
bench.add('writeSheetXml - auto-detect widths', async () => {
  const start = performance.now();
  const rows = generateRows();
  const options: WriteSheetXmlOptions = {
    columnWidths: {
      autoDetectColumnWidth: true,
    },
  };
  const xmlStream = writeSheetXml(rows, options);
  const bytes = await consumeXml(xmlStream);
  const time = performance.now() - start;

  assert(bytes > 0, 'Should generate XML');
  return time;
});

async function main() {
  console.log(`\nBenchmarking writeSheetXml with ${ROW_COUNT.toLocaleString()} rows, ${COLUMN_COUNT} columns\n`);

  await bench.run();

  console.log('\n=== Benchmark Results ===');
  console.log(`Rows: ${ROW_COUNT.toLocaleString()}`);
  console.log(`Columns: ${COLUMN_COUNT}`);
  console.log('\nSummary:');
  for (let i = 0; i < bench.results.length; i++) {
    const result = bench.results[i];
    assert(result, 'Benchmark should have result');
    assert(result.mean > 0, 'Benchmark should take time');
    const task = bench.tasks[i];
    const name = task?.name || `Benchmark ${i + 1}`;
    console.log(`  ${name}: ${result.mean.toFixed(2)}ms (min: ${result.min.toFixed(2)}ms, max: ${result.max.toFixed(2)}ms)`);
  }

  console.log('\n✓ All benchmarks completed successfully!');
}

main().catch((error) => {
  console.error('Benchmark failed:', error);
  process.exit(1);
});
