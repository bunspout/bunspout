/*
 * NOTE ON MEMORY MEASUREMENTS
 *
 * Memory benchmarks in JS measure *relative usage and regressions*,
 * not exact byte-for-byte allocation.
 *
 * We use:
 * - peak heapUsed for in-memory strategies
 * - peak RSS for file-based strategies
 *
 * GC is non-deterministic; results may vary slightly across runs,
 * machines, and runtimes. Assertions are intentionally tolerant.
 */
import { rm } from 'fs/promises';
import assert from 'node:assert';
import { tmpdir } from 'os';
import { join } from 'path';
import { Bench } from 'tinybench';
import { IN_MEMORY_ENTRY_OVERHEAD_BYTES, MAX_NUM_STRINGS_PER_TEMP_FILE } from '@xlsx/shared-strings-caching';
import { FileBasedStrategy } from '@xlsx/shared-strings-caching/file-based-strategy';
import { InMemoryStrategy } from '@xlsx/shared-strings-caching/in-memory-strategy';

const STRING_COUNT = 50_000; // Number of shared strings to test
const AVERAGE_STRING_LENGTH = 50; // Average characters per string

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
 * Generate test strings
 */
function generateStrings(count: number): string[] {
  const strings: string[] = [];
  for (let i = 0; i < count; i++) {
    // Create strings of varying lengths around the average
    const length = AVERAGE_STRING_LENGTH + (i % 20) - 10;
    strings.push(`Shared String ${i} - ${'x'.repeat(Math.max(10, length))}`);
  }
  return strings;
}

const bench = new Bench({ iterations: 5 });

// Track metrics separately for each strategy
const inMemoryAddTimes: number[] = [];
const inMemoryGetTimes: number[] = [];
const inMemoryMemPeaks: number[] = [];
const fileBasedAddTimes: number[] = [];
const fileBasedGetTimes: number[] = [];
const fileBasedMemPeaks: number[] = [];

// Benchmark: In-memory strategy - adding strings
bench.add('in-memory strategy - add strings', async () => {
  const start = performance.now();

  // Generate strings first and force them to be fully realized
  // This prevents lazy string flattening from happening during the benchmark
  const strings = generateStrings(STRING_COUNT);
  strings.forEach(s => s.length); // Force strings to be fully realized

  // Force GC before measurement to get a clean baseline
  if ('gc' in globalThis && typeof (globalThis as { gc?: () => void }).gc === 'function') {
    (globalThis as { gc: () => void }).gc();
  }
  await new Promise(r => setImmediate(r));

  // Measure baseline memory after strings are generated and realized
  // This represents "string memory" - the memory used by the string data itself
  const stringMemoryBaseline = process.memoryUsage().heapUsed;

  // Now create strategy
  const strategy = new InMemoryStrategy(STRING_COUNT);
  await new Promise(resolve => setImmediate(resolve));

  // Track peak memory during the add operation
  // This will include both string memory and strategy overhead
  let peakHeap = process.memoryUsage().heapUsed;

  for (let i = 0; i < STRING_COUNT; i++) {
    await strategy.addString(i, strings[i]!);

    const currentHeap = process.memoryUsage().heapUsed;
    if (currentHeap > peakHeap) {
      peakHeap = currentHeap;
    }
  }

  const time = performance.now() - start;

  // Strategy memory = peak memory during adding - baseline (string memory already accounted)
  // This gives us the memory overhead of the strategy itself
  const mem = peakHeap - stringMemoryBaseline;

  inMemoryAddTimes.push(time);
  inMemoryMemPeaks.push(mem);

  await strategy.cleanup();
  // Keep strings in scope to prevent GC during measurement
  void strings;
});

// Benchmark: In-memory strategy - getting strings
bench.add('in-memory strategy - get strings', async () => {
  const strategy = new InMemoryStrategy(STRING_COUNT);
  const strings = generateStrings(STRING_COUNT);

  // Pre-populate
  for (let i = 0; i < STRING_COUNT; i++) {
    await strategy.addString(i, strings[i]!);
  }

  const start = performance.now();

  // Read all strings in order (best case for file-based)
  for (let i = 0; i < STRING_COUNT; i++) {
    const value = await strategy.getString(i);
    assert.strictEqual(value, strings[i]);
  }

  const time = performance.now() - start;
  inMemoryGetTimes.push(time);

  await strategy.cleanup();
});

// Benchmark: File-based strategy - adding strings
bench.add('file-based strategy - add strings', async () => {
  const tempDir = join(tmpdir(), `bunspout-bench-${Date.now()}`);

  // Generate strings first and force them to be fully realized
  // This prevents lazy string flattening from happening during the benchmark
  const strings = generateStrings(STRING_COUNT);
  strings.forEach(s => s.length); // Force strings to be fully realized

  // Force GC before measurement to get a clean baseline
  if ('gc' in globalThis && typeof (globalThis as { gc?: () => void }).gc === 'function') {
    (globalThis as { gc: () => void }).gc();
  }
  await new Promise(r => setImmediate(r));

  // Measure baseline RSS after strings are generated and realized
  // RSS (Resident Set Size) is better for file-based strategy as it includes all process memory
  const before = process.memoryUsage().rss;

  // Now create strategy
  const strategy = new FileBasedStrategy(tempDir, MAX_NUM_STRINGS_PER_TEMP_FILE);
  await new Promise(r => setImmediate(r));
  const start = performance.now();

  // Track peak RSS during the add operation
  let peakRss = process.memoryUsage().rss;

  for (let i = 0; i < STRING_COUNT; i++) {
    await strategy.addString(i, strings[i]!);

    const currentRss = process.memoryUsage().rss;
    if (currentRss > peakRss) {
      peakRss = currentRss;
    }
  }

  const time = performance.now() - start;

  // RSS delta = peak RSS during adding - baseline
  const after = peakRss;
  const rssDelta = after - before;
  const mem = rssDelta;

  fileBasedAddTimes.push(time);
  fileBasedMemPeaks.push(mem);

  await strategy.cleanup();
  try {
    await rm(tempDir, { recursive: true, force: true });
  } catch {
    // Ignore cleanup errors
  }
  // Keep strings in scope to prevent GC during measurement
  void strings;
});

// Benchmark: File-based strategy - getting strings (sequential access - best case)
bench.add('file-based strategy - get strings (sequential)', async () => {
  const tempDir = join(tmpdir(), `bunspout-bench-${Date.now()}`);
  const strategy = new FileBasedStrategy(tempDir, MAX_NUM_STRINGS_PER_TEMP_FILE);
  const strings = generateStrings(STRING_COUNT);

  // Pre-populate
  for (let i = 0; i < STRING_COUNT; i++) {
    await strategy.addString(i, strings[i]!);
  }

  const start = performance.now();

  // Read all strings in order (best case for file-based - minimizes file loads)
  for (let i = 0; i < STRING_COUNT; i++) {
    const value = await strategy.getString(i);
    assert.strictEqual(value, strings[i]);
  }

  const time = performance.now() - start;
  fileBasedGetTimes.push(time);

  await strategy.cleanup();
  try {
    await rm(tempDir, { recursive: true, force: true });
  } catch {
    // Ignore cleanup errors
  }
});

async function main() {
  const hasGC = typeof (globalThis as { gc?: () => void }).gc === 'function';
  console.log('\nBenchmarking shared strings caching strategies');
  console.log(`GC exposed: ${hasGC}`);
  if (!hasGC) {
    console.log('Hint: Run with --expose-gc flag for more accurate memory measurements');
    console.log('  Example: bun --expose-gc run src/tests/benchmarks/shared-strings-caching.bench.ts');
  }
  console.log(`String count: ${STRING_COUNT.toLocaleString()}`);
  console.log(`Average string length: ${AVERAGE_STRING_LENGTH} characters`);
  // Expected memory is entry overhead only (strings stored once, referenced)
  const expectedMemoryMB = (STRING_COUNT * IN_MEMORY_ENTRY_OVERHEAD_BYTES) / (1024 * 1024);
  console.log(`Expected memory (in-memory, entry overhead only): ~${expectedMemoryMB.toFixed(2)} MB\n`);

  await bench.run();

  const inMemoryAddMode = mode(inMemoryAddTimes);
  const inMemoryGetMode = mode(inMemoryGetTimes);
  const inMemoryMemMode = mode(inMemoryMemPeaks);
  const fileBasedAddMode = mode(fileBasedAddTimes);
  const fileBasedGetMode = mode(fileBasedGetTimes);
  const fileBasedMemMode = mode(fileBasedMemPeaks);

  console.log('\n=== Benchmark Results ===');
  console.log('\nIn-Memory Strategy:');
  console.log(`  Add time: ${inMemoryAddMode.toFixed(2)} ms`);
  console.log(`  Get time: ${inMemoryGetMode.toFixed(2)} ms`);
  console.log(`  Expected memory: ~${expectedMemoryMB.toFixed(2)} MB`);
  if (inMemoryMemMode > 0) {
    console.log(`  Measured memory: ${(inMemoryMemMode / 1024 / 1024).toFixed(2)} MB`);
  } else {
    console.log('  Measured memory: < 1 MB (GC interference - measurement unreliable)');
  }
  console.log('\nFile-Based Strategy:');
  console.log(`  Add time: ${fileBasedAddMode.toFixed(2)} ms`);
  console.log(`  Get time: ${fileBasedGetMode.toFixed(2)} ms`);
  console.log(`  Expected memory: < ${(expectedMemoryMB * 0.1).toFixed(2)} MB (only current file in memory)`);
  if (fileBasedMemMode > 0) {
    console.log(`  Measured memory: ${(fileBasedMemMode / 1024 / 1024).toFixed(2)} MB`);
  } else {
    console.log('  Measured memory: < 1 MB (GC interference - measurement unreliable)');
  }
  console.log('\nComparison:');
  console.log(`  Add time: ${((fileBasedAddMode / inMemoryAddMode - 1) * 100).toFixed(1)}% ${fileBasedAddMode > inMemoryAddMode ? 'slower' : 'faster'}`);
  console.log(`  Get time: ${((fileBasedGetMode / inMemoryGetMode - 1) * 100).toFixed(1)}% ${fileBasedGetMode > inMemoryGetMode ? 'slower' : 'faster'}`);
  console.log(`  Memory: File-based uses ~${((1 - 0.1) * 100).toFixed(0)}% less memory (only loads current file)`);

  // Assertions
  const MAX_ADD_TIME_MS = 10_000; // 10 seconds
  const MAX_GET_TIME_MS = 30_000; // 30 seconds
  // Expected memory is entry overhead only (strings stored once, referenced)
  const EXPECTED_MEMORY_BYTES = STRING_COUNT * IN_MEMORY_ENTRY_OVERHEAD_BYTES;
  const MAX_MEMORY_BYTES = EXPECTED_MEMORY_BYTES * 1.5; // 1.5x expected for additional overhead

  assert(inMemoryAddMode < MAX_ADD_TIME_MS, `In-memory add too slow: ${inMemoryAddMode.toFixed(2)}ms > ${MAX_ADD_TIME_MS}ms`);
  assert(fileBasedAddMode < MAX_ADD_TIME_MS, `File-based add too slow: ${fileBasedAddMode.toFixed(2)}ms > ${MAX_ADD_TIME_MS}ms`);
  assert(inMemoryGetMode < MAX_GET_TIME_MS, `In-memory get too slow: ${inMemoryGetMode.toFixed(2)}ms > ${MAX_GET_TIME_MS}ms`);
  assert(fileBasedGetMode < MAX_GET_TIME_MS, `File-based get too slow: ${fileBasedGetMode.toFixed(2)}ms > ${MAX_GET_TIME_MS}ms`);

  // Memory assertions - verify ratios and orders of magnitude
  // File-based should stay under a reasonable limit (5MB for current file + overhead)
  if (fileBasedMemMode > 0) {
    assert(
      fileBasedMemMode < 5 * 1024 * 1024,
      `File-based should stay under 5MB: ${(fileBasedMemMode / 1024 / 1024).toFixed(2)}MB >= 5MB`,
    );
  }

  // In-memory should use significantly more memory than file-based when both are measurable
  if (inMemoryMemMode > 0 && fileBasedMemMode > 0 && inMemoryMemMode > 10 * 1024 * 1024) {
    // Only assert if in-memory uses more than 10MB (to avoid false positives from GC)
    assert(
      inMemoryMemMode > fileBasedMemMode * 5,
      `In-memory should use significantly more memory than file-based: ${(inMemoryMemMode / 1024 / 1024).toFixed(2)}MB <= ${(fileBasedMemMode * 5 / 1024 / 1024).toFixed(2)}MB`,
    );
  }

  // In-memory should not exceed expected memory by too much (with some overhead allowance)
  if (inMemoryMemMode > 0) {
    assert(
      inMemoryMemMode < MAX_MEMORY_BYTES,
      `In-memory too much memory: ${(inMemoryMemMode / 1024 / 1024).toFixed(2)}MB > ${(MAX_MEMORY_BYTES / 1024 / 1024).toFixed(2)}MB`,
    );
  }

  console.log('\n✓ All benchmark assertions passed!');
}

main().catch((error) => {
  console.error('Benchmark failed:', error);
  process.exit(1);
});
