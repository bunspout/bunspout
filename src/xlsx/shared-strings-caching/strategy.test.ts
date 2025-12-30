import { rm } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';
import { describe, test, expect, beforeEach, afterEach } from 'bun:test';
import { CachingStrategyFactory, IN_MEMORY_ENTRY_OVERHEAD_BYTES, MAX_NUM_STRINGS_PER_TEMP_FILE } from './factory';
import { FileBasedStrategy } from './file-based-strategy';
import { InMemoryStrategy } from './in-memory-strategy';

describe('InMemoryStrategy', () => {
  let strategy: InMemoryStrategy;

  beforeEach(() => {
    strategy = new InMemoryStrategy();
  });

  afterEach(async () => {
    await strategy.cleanup();
  });

  test('should add and retrieve strings', async () => {
    await strategy.addString(0, 'Hello');
    await strategy.addString(1, 'World');

    expect(await strategy.getString(0)).toBe('Hello');
    expect(await strategy.getString(1)).toBe('World');
    expect(strategy.getCount()).toBe(2);
  });

  test('should handle sparse indices', async () => {
    await strategy.addString(0, 'First');
    await strategy.addString(5, 'Sixth');

    expect(await strategy.getString(0)).toBe('First');
    expect(await strategy.getString(5)).toBe('Sixth');
    expect(await strategy.getString(3)).toBeUndefined();
    expect(strategy.getCount()).toBe(6);
  });

  test('should handle initial capacity', async () => {
    const strategyWithCapacity = new InMemoryStrategy(100);
    await strategyWithCapacity.addString(50, 'Middle');
    expect(await strategyWithCapacity.getString(50)).toBe('Middle');
    await strategyWithCapacity.cleanup();
  });
});

describe('FileBasedStrategy', () => {
  let strategy: FileBasedStrategy;
  let tempDir: string;

  beforeEach(() => {
    tempDir = join(tmpdir(), `bunspout-test-${Date.now()}`);
    strategy = new FileBasedStrategy(tempDir, 10); // Small max per file for testing
  });

  afterEach(async () => {
    await strategy.cleanup();
    try {
      await rm(tempDir, { recursive: true, force: true });
    } catch {
      // Ignore cleanup errors
    }
  });

  test('should add and retrieve strings', async () => {
    await strategy.addString(0, 'Hello');
    await strategy.addString(1, 'World');

    expect(await strategy.getString(0)).toBe('Hello');
    expect(await strategy.getString(1)).toBe('World');
    expect(strategy.getCount()).toBe(2);
  });

  test('should handle multiple files', async () => {
    // Add strings that span multiple files (max 10 per file)
    for (let i = 0; i < 25; i++) {
      await strategy.addString(i, `String${i}`);
    }

    // Retrieve strings from different files
    expect(await strategy.getString(0)).toBe('String0');
    expect(await strategy.getString(9)).toBe('String9');
    expect(await strategy.getString(10)).toBe('String10');
    expect(await strategy.getString(24)).toBe('String24');
    expect(strategy.getCount()).toBe(25);
  });

  test('should handle sparse indices', async () => {
    await strategy.addString(0, 'First');
    await strategy.addString(15, 'Sixteenth');

    expect(await strategy.getString(0)).toBe('First');
    expect(await strategy.getString(15)).toBe('Sixteenth');
    expect(await strategy.getString(10)).toBeUndefined();
  });

  test('should handle out-of-order additions (backwards)', async () => {
    // Add to a later file first (fileIndex 1)
    await strategy.addString(15, 'Later');
    // Then add to an earlier file (fileIndex 0) - this was the bug case
    await strategy.addString(5, 'Earlier');

    expect(await strategy.getString(5)).toBe('Earlier');
    expect(await strategy.getString(15)).toBe('Later');
    expect(strategy.getCount()).toBe(16);
  });
});

describe('CachingStrategyFactory', () => {
  test('should create in-memory strategy for small counts', async () => {
    // Mock memory limit to return a large value
    const mockGetMemoryLimit = () => 1000000; // 1GB in KB
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    const strategy = factory.createBestCachingStrategy(100);
    expect(strategy).toBeInstanceOf(InMemoryStrategy);
    await strategy.cleanup();
  });

  test('should create file-based strategy when count is unknown', async () => {
    const mockGetMemoryLimit = () => 1000000;
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    const strategy = factory.createBestCachingStrategy(null);
    expect(strategy).toBeInstanceOf(FileBasedStrategy);
    await strategy.cleanup();
  });

  test('should use file-based strategy for very large counts', async () => {
    // Mock memory limit to return a small value that won't fit the large count
    const largeCount = 1000000;
    // Calculate memory needed: entry overhead in bytes, converted to KB
    const memoryNeededKB = Math.ceil((largeCount * IN_MEMORY_ENTRY_OVERHEAD_BYTES) / 1024);
    const mockGetMemoryLimit = () => memoryNeededKB - 1; // Just below what's needed
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    const strategy = factory.createBestCachingStrategy(largeCount);
    expect(strategy).toBeInstanceOf(FileBasedStrategy);
    await strategy.cleanup();
  });

  test('should use in-memory strategy when memory is sufficient', async () => {
    const count = 1000;
    // Calculate memory needed: entry overhead in bytes, converted to KB
    // With 50% safety margin, we need heap limit >= 2 * memory needed
    const memoryNeededBytes = count * IN_MEMORY_ENTRY_OVERHEAD_BYTES;
    const memoryNeededKB = Math.ceil(memoryNeededBytes / 1024);
    // Heap limit must be at least 2x memory needed (50% safety margin)
    const mockGetMemoryLimit = () => memoryNeededKB * 2 + 1000; // More than needed with margin
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    const strategy = factory.createBestCachingStrategy(count);
    expect(strategy).toBeInstanceOf(InMemoryStrategy);
    await strategy.cleanup();
  });

  test('should use file-based strategy when memory limit is very small', async () => {
    // Simulate a very small heap limit that won't fit the count
    const mockGetMemoryLimit = () => 100; // 100 KB - very small
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    // With very small memory, should use file-based
    const strategy = factory.createBestCachingStrategy(MAX_NUM_STRINGS_PER_TEMP_FILE);
    expect(strategy).toBeInstanceOf(FileBasedStrategy);
    await strategy.cleanup();
  });

  test('should use file-based strategy when memory limit is insufficient', async () => {
    // Simulate a heap limit that's too small for the count (even with safety margin)
    const count = 1000;
    const memoryNeededBytes = count * IN_MEMORY_ENTRY_OVERHEAD_BYTES;
    const memoryNeededKB = Math.ceil(memoryNeededBytes / 1024);
    // Heap limit must be at least 2x memory needed (50% safety margin)
    // Use a value that's less than 2x, so it should use file-based
    const mockGetMemoryLimit = () => memoryNeededKB * 1.5; // Less than 2x needed
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    // Should use file-based when memory is insufficient
    const strategy = factory.createBestCachingStrategy(count);
    expect(strategy).toBeInstanceOf(FileBasedStrategy);
    await strategy.cleanup();
  });

  test('static method should work with default memory limit', async () => {
    const strategy = CachingStrategyFactory.createBestCachingStrategy(100);
    expect(strategy).toBeDefined();
    await strategy.cleanup();
  });

  test('should reject negative counts and use file-based strategy', async () => {
    const mockGetMemoryLimit = () => 1000000;
    const factory = new CachingStrategyFactory(mockGetMemoryLimit);
    // Negative count should be rejected and fall back to file-based strategy
    const strategy = factory.createBestCachingStrategy(-1);
    expect(strategy).toBeInstanceOf(FileBasedStrategy);
    await strategy.cleanup();
  });

  test('should export constants', () => {
    expect(IN_MEMORY_ENTRY_OVERHEAD_BYTES).toBe(48);
    expect(MAX_NUM_STRINGS_PER_TEMP_FILE).toBe(10000);
  });
});
