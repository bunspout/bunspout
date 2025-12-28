/**
 * Interface for shared strings caching strategies
 */
export interface SharedStringsCachingStrategy {
  /**
   * Adds a string to the cache at the given index
   */
  addString(index: number, value: string): Promise<void>;

  /**
   * Gets a string from the cache by index
   */
  getString(index: number): Promise<string | undefined>;

  /**
   * Gets the total count of strings in the cache
   */
  getCount(): number;

  /**
   * Cleans up any temporary resources (files, etc.)
   */
  cleanup(): Promise<void>;
}

