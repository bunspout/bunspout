import { randomUUID } from 'crypto';
import { tmpdir } from 'os';
import { join } from 'path';
import { FileBasedStrategy } from './file-based-strategy';
import { InMemoryStrategy } from './in-memory-strategy';
import { getMemoryLimitInKB } from './memory-limit';
import type { SharedStringsCachingStrategy } from './strategy';

/**
 * Per-entry memory overhead for in-memory shared strings strategy (in bytes)
 *
 * In JavaScript, strings are stored once in the heap and referenced by the array.
 * This constant estimates the overhead per array entry, not the string payload size.
 *
 * Overhead includes:
 * - Array slot (pointer/reference): ~8 bytes
 * - Array metadata and growth overhead: ~24-32 bytes
 * - Conservative estimate: 48 bytes per entry
 */
export const IN_MEMORY_ENTRY_OVERHEAD_BYTES = 48;

/**
 * To avoid running out of memory when extracting a huge number of shared strings, they can be saved to temporary files
 * instead of in memory. Then, when accessing a string, the corresponding file contents will be loaded in memory
 * and the string will be quickly retrieved.
 * The performance bottleneck is not when creating these temporary files, but rather when loading their content.
 * Because the contents of the last loaded file stays in memory until another file needs to be loaded, it works
 * best when the indexes of the shared strings are sorted in the sheet data.
 * 10,000 was chosen because it creates small files that are fast to be loaded in memory.
 */
export const MAX_NUM_STRINGS_PER_TEMP_FILE = 10000;

/**
 * Factory for creating the best caching strategy based on available memory
 */
export class CachingStrategyFactory {
  private getMemoryLimit: () => number;

  constructor(getMemoryLimit: () => number = getMemoryLimitInKB) {
    this.getMemoryLimit = getMemoryLimit;
  }

  /**
   * Creates the best caching strategy, given the number of unique shared strings
   * and the heap size limit.
   *
   * In JavaScript, strings are stored once in the heap and referenced by the array.
   * Memory estimation uses per-entry overhead, not string payload size.
   *
   * @param sharedStringsUniqueCount Number of unique shared strings (null if unknown)
   * @param tempFolder Optional temporary folder where the temporary files will be stored
   * @returns The best caching strategy
   */
  createBestCachingStrategy(
    sharedStringsUniqueCount: number | null,
    tempFolder?: string,
  ): SharedStringsCachingStrategy {
    if (this.isInMemoryStrategyUsageSafe(sharedStringsUniqueCount)) {
      return new InMemoryStrategy(sharedStringsUniqueCount ?? undefined);
    }

    // Use file-based strategy
    const finalTempFolder =
      tempFolder ?? join(tmpdir(), 'bunspout-shared-strings', randomUUID());
    return new FileBasedStrategy(finalTempFolder, MAX_NUM_STRINGS_PER_TEMP_FILE);
  }

  /**
   * Returns whether it is safe to use in-memory caching, given the number of unique shared strings
   * and the heap size limit.
   *
   * Safety criteria:
   * - Reject when count is unknown (null)
   * - Apply hard upper bound (MAX_NUM_STRINGS_PER_TEMP_FILE)
   * - Estimate memory using reference overhead, not string payload size
   * - Use safety margin (≤50% of heap limit) to account for GC variance
   * - Default to conservative behavior when limits are unknown
   *
   * @param sharedStringsUniqueCount Number of unique shared strings (null if unknown)
   */
  private isInMemoryStrategyUsageSafe(
    sharedStringsUniqueCount: number | null,
  ): boolean {
    if (sharedStringsUniqueCount === null || sharedStringsUniqueCount < 0) {
      return false;
    }

    // Apply hard upper bound: never use in-memory for counts exceeding temp file threshold
    if (sharedStringsUniqueCount >= MAX_NUM_STRINGS_PER_TEMP_FILE) {
      return false;
    }

    const heapLimitKB = this.getMemoryLimit();

    // getMemoryLimit() always returns a value (conservative hard cap if actual limit unknown)
    // Convert heap limit from KB to bytes
    const heapLimitBytes = heapLimitKB * 1024;

    // Estimate memory needed: entry overhead only (strings are stored once, referenced)
    const memoryNeededBytes = sharedStringsUniqueCount * IN_MEMORY_ENTRY_OVERHEAD_BYTES;

    // Use safety margin: only use in-memory if needed memory ≤ 50% of heap limit
    // This accounts for GC variance and other heap allocations
    const safetyMargin = 0.5;
    const maxSafeMemoryBytes = heapLimitBytes * safetyMargin;

    return memoryNeededBytes <= maxSafeMemoryBytes;
  }

  /**
   * Static convenience method that uses the default memory limit function
   */
  static createBestCachingStrategy(
    sharedStringsUniqueCount: number | null,
    tempFolder?: string,
  ): SharedStringsCachingStrategy {
    const factory = new CachingStrategyFactory();
    return factory.createBestCachingStrategy(sharedStringsUniqueCount, tempFolder);
  }
}
