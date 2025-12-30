import { describe, test, expect } from 'bun:test';
import {
  stringToBytes,
  bytesToString,
  bytesToUint8Array,
  uint8ArrayToBytes,
} from './common';

describe('Common Adapters', () => {
  describe('stringToBytes', () => {
    test('should convert string to AsyncIterable<Uint8Array>', async () => {
      const result: Uint8Array[] = [];
      for await (const chunk of stringToBytes('Hello World')) {
        result.push(chunk);
      }
      const text = new TextDecoder().decode(
        new Uint8Array(result.reduce((acc, chunk) => {
          const combined = new Uint8Array(acc.length + chunk.length);
          combined.set(acc);
          combined.set(chunk, acc.length);
          return combined;
        }, new Uint8Array(0))),
      );
      expect(text).toBe('Hello World');
    });
  });

  describe('bytesToString', () => {
    test('should convert AsyncIterable<Uint8Array> to string', async () => {
      const bytes = async function* () {
        yield new TextEncoder().encode('Hello');
        yield new TextEncoder().encode(' ');
        yield new TextEncoder().encode('World');
      }();
      const result = await bytesToString(bytes);
      expect(result).toBe('Hello World');
    });
  });

  describe('bytesToUint8Array', () => {
    test('should convert AsyncIterable<Uint8Array> to single Uint8Array', async () => {
      const bytes = async function* () {
        yield new Uint8Array([1, 2, 3]);
        yield new Uint8Array([4, 5, 6]);
      }();
      const result = await bytesToUint8Array(bytes);
      expect(result).toEqual(new Uint8Array([1, 2, 3, 4, 5, 6]));
    });
  });

  describe('uint8ArrayToBytes', () => {
    test('should convert Uint8Array to AsyncIterable<Uint8Array>', async () => {
      const data = new Uint8Array([1, 2, 3, 4, 5]);
      const result: Uint8Array[] = [];
      for await (const chunk of uint8ArrayToBytes(data)) {
        result.push(chunk);
      }
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual(data);
    });
  });
});
