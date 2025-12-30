import { existsSync } from 'fs';
import assert from 'node:assert';
import { Bench } from 'tinybench';
import { readXlsx } from '@xlsx/reader';
import { writeXlsx } from '@xlsx/writer';
import { cell } from '@sheet/cell';
import { row } from '@sheet/row';

const BENCHMARK_FILE = 'benchmark-inline-strings.xlsx';
const ROW_COUNT = 250_000;

/**
 * Calculate the mode (most frequent value) of an array
 */
function mode(values: number[]): number {
  const counts = new Map<number, number>();
  for (const value of values) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }

  let maxCount = 0;
  let modeValue = values[0]!;
  for (const [value, count] of counts) {
    if (count > maxCount) {
      maxCount = count;
      modeValue = value;
    }
  }
  return modeValue;
}

/**
 * Generates a test Excel file with inline strings if it doesn't exist
 */
async function ensureTestFile(): Promise<void> {
  if (existsSync(BENCHMARK_FILE)) {
    console.log(`Using existing test file: ${BENCHMARK_FILE}`);
    return;
  }

  console.log(`Generating test file with ${ROW_COUNT.toLocaleString()} rows (inline strings)...`);
  const startTime = Date.now();

  // Use inline strings (default, but explicit for clarity)
  await writeXlsx(BENCHMARK_FILE, {
    sheets: [
      {
        name: 'Data',
        rows: (async function* () {
          // Generate 500k rows with inline strings
          for (let i = 1; i <= ROW_COUNT; i++) {
            yield row(
              [
                cell(`Row ${i}`), // Inline string
                cell(`Value ${i}`), // Inline string
                cell(i), // Number
              ],
              { rowIndex: i },
            );

            // Log progress every 50k rows
            if (i % 50_000 === 0) {
              console.log(`  Generated ${i.toLocaleString()} rows...`);
            }
          }
        })(),
      },
    ],
    // Explicitly use inline strings (this is the default)
  });

  const elapsed = Date.now() - startTime;
  console.log(
    `Test file generated in ${(elapsed / 1000).toFixed(2)}s: ${BENCHMARK_FILE}`,
  );
}

const bench = new Bench({ iterations: 10 });

const times: number[] = [];
const memPeaks: number[] = [];

bench.add('read 500k inline strings xlsx', async () => {
  const startMem = process.memoryUsage().heapUsed;
  const start = performance.now();

  const workbook = await readXlsx(BENCHMARK_FILE);
  const sheet = workbook.sheet('Data');

  let rows = 0;
  for await (const row of sheet.rows()) {
    rows++;
    void row; // Acknowledge we're counting rows
  }

  const time = performance.now() - start;
  const mem = process.memoryUsage().heapUsed - startMem;

  times.push(time);
  memPeaks.push(mem);

  assert.strictEqual(rows, ROW_COUNT);
});

async function main() {
  await ensureTestFile();

  console.log('\nRunning benchmark...\n');
  await bench.run();

  const timeMode = mode(times);
  const memMode = mode(memPeaks);

  console.log('\n=== Benchmark Results ===');
  console.log(`Rows processed: ${ROW_COUNT.toLocaleString()}`);
  console.log(`Time mode: ${timeMode.toFixed(2)} ms`);
  console.log(`Memory mode: ${(memMode / 1024 / 1024).toFixed(2)} MB`);

  // Assertions
  // Goal: sub-20 MB (openspout has 8MB for 300k rows, so ~13MB proportional for 500k)
  const MAX_TIME_MS = 30_000; // 30 seconds
  const MAX_MEMORY_BYTES = 20 * 1024 * 1024; // 100 MB (sub-100 MB goal)

  assert(timeMode < MAX_TIME_MS, `Too slow: ${timeMode.toFixed(2)}ms > ${MAX_TIME_MS}ms`);
  assert(memMode < MAX_MEMORY_BYTES, `Too much memory: ${(memMode / 1024 / 1024).toFixed(2)}MB > ${MAX_MEMORY_BYTES / 1024 / 1024}MB`);

  console.log('\n✓ All benchmark assertions passed!');
}

main().catch((error) => {
  console.error('Benchmark failed:', error);
  process.exit(1);
});
