/**
 * Gets the heap size limit in KB
 * Returns a conservative hard cap if the limit cannot be determined
 *
 * In JavaScript, strings are stored once in the heap and referenced.
 * This function returns the actual heap size limit when available, or a conservative fallback.
 */
export function getMemoryLimitInKB(): number {
  if (typeof process === 'undefined') {
    // Use conservative hard cap when process is unavailable
    const CONSERVATIVE_HEAP_LIMIT_KB = 512 * 1024; // 512 MB
    return CONSERVATIVE_HEAP_LIMIT_KB;
  }

  // On Node.js, try to get the actual heap size limit from v8
  // Use synchronous require since this is called synchronously
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const v8 = require('v8');
    if (v8 && typeof v8.getHeapStatistics === 'function') {
      const stats = v8.getHeapStatistics();
      const heapSizeLimit = stats.heap_size_limit;
      if (heapSizeLimit && heapSizeLimit > 0) {
        // Convert bytes to KB
        return Math.floor(heapSizeLimit / 1024);
      }
    }
  } catch {
    // v8 module not available (e.g., Bun or older Node.js)
  }

  // Fallback: Use conservative hard cap
  // Default to 512MB heap limit for safety when we can't detect the actual limit
  // This prevents OOM under GC variance
  const CONSERVATIVE_HEAP_LIMIT_KB = 512 * 1024; // 512 MB
  return CONSERVATIVE_HEAP_LIMIT_KB;
}

