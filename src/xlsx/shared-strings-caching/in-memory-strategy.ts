import type { SharedStringsCachingStrategy } from './strategy';

/**
 * In-memory caching strategy for shared strings
 * Stores all strings in a simple array
 */
export class InMemoryStrategy implements SharedStringsCachingStrategy {
  private strings: string[] = [];
  private count: number = 0;

  constructor(initialCapacity?: number) {
    if (initialCapacity !== undefined) {
      this.strings = new Array(initialCapacity);
    }
  }

  async addString(index: number, value: string): Promise<void> {
    // Ensure array is large enough
    if (index >= this.strings.length) {
      // Grow array to accommodate the index
      const newLength = Math.max(index + 1, this.strings.length * 2);
      const newArray = new Array(newLength);
      for (let i = 0; i < this.strings.length; i++) {
        newArray[i] = this.strings[i];
      }
      this.strings = newArray;
    }

    this.strings[index] = value;
    this.count = Math.max(this.count, index + 1);
  }

  async getString(index: number): Promise<string | undefined> {
    return this.strings[index];
  }

  getCount(): number {
    return this.count;
  }

  async cleanup(): Promise<void> {
    // No cleanup needed for in-memory strategy
    this.strings = [];
    this.count = 0;
  }
}

