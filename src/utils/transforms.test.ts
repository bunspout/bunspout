import { describe, test, expect } from 'bun:test';
import { mapRows, filterRows, limitRows, collect } from './transforms';
import type { Row } from '../types';

describe('Row Transforms', () => {
  describe('mapRows', () => {
    test('should map rows synchronously', async () => {
      const rows: Row[] = [
        { cells: [{ value: 'A' }] },
        { cells: [{ value: 'B' }] },
      ];
      const result: Row[] = [];
      for await (const row of mapRows(async function* () {
        for (const r of rows) yield r;
      }(), (r) => ({
        ...r,
        cells: r.cells
          .filter((c): c is NonNullable<typeof c> => c != null)
          .map((c) => ({ ...c, value: String(c.value).toUpperCase() })),
      }))) {
        result.push(row);
      }
      expect(result).toHaveLength(2);
      expect(result[0]?.cells[0]?.value).toBe('A');
      expect(result[1]?.cells[0]?.value).toBe('B');
    });

    test('should map rows asynchronously', async () => {
      const rows: Row[] = [{ cells: [{ value: 1 }] }, { cells: [{ value: 2 }] }];
      const result: Row[] = [];
      for await (const row of mapRows(async function* () {
        for (const r of rows) yield r;
      }(), async (r) => ({
        ...r,
        cells: r.cells
          .filter((c): c is NonNullable<typeof c> => c != null)
          .map((c) => ({ ...c, value: Number(c.value) * 2 })),
      }))) {
        result.push(row);
      }
      expect(result).toHaveLength(2);
      expect(result[0]?.cells[0]?.value).toBe(2);
      expect(result[1]?.cells[0]?.value).toBe(4);
    });
  });

  describe('filterRows', () => {
    test('should filter rows', async () => {
      const rows: Row[] = [
        { cells: [{ value: 'A' }] },
        { cells: [{ value: 'B' }] },
        { cells: [{ value: 'C' }] },
      ];
      const result: Row[] = [];
      for await (const row of filterRows(async function* () {
        for (const r of rows) yield r;
      }(), (r) => r.cells[0]?.value === 'B')) {
        result.push(row);
      }
      expect(result).toHaveLength(1);
      expect(result[0]?.cells[0]?.value).toBe('B');
    });

    test('should handle empty filter result', async () => {
      const rows: Row[] = [{ cells: [{ value: 'A' }] }];
      const result: Row[] = [];
      for await (const row of filterRows(async function* () {
        for (const r of rows) yield r;
      }(), () => false)) {
        result.push(row);
      }
      expect(result).toHaveLength(0);
    });
  });

  describe('limitRows', () => {
    test('should limit number of rows', async () => {
      const rows: Row[] = [
        { cells: [{ value: 'A' }] },
        { cells: [{ value: 'B' }] },
        { cells: [{ value: 'C' }] },
        { cells: [{ value: 'D' }] },
      ];
      const result: Row[] = [];
      for await (const row of limitRows(async function* () {
        for (const r of rows) yield r;
      }(), 2)) {
        result.push(row);
      }
      expect(result).toHaveLength(2);
      expect(result[0]?.cells[0]?.value).toBe('A');
      expect(result[1]?.cells[0]?.value).toBe('B');
    });

    test('should handle limit greater than available rows', async () => {
      const rows: Row[] = [{ cells: [{ value: 'A' }] }];
      const result: Row[] = [];
      for await (const row of limitRows(async function* () {
        for (const r of rows) yield r;
      }(), 10)) {
        result.push(row);
      }
      expect(result).toHaveLength(1);
    });
  });

  describe('collect', () => {
    test('should collect all items from async iterable', async () => {
      const items = [1, 2, 3, 4, 5];
      const result = await collect(async function* () {
        for (const item of items) yield item;
      }());
      expect(result).toEqual([1, 2, 3, 4, 5]);
    });

    test('should handle empty iterable', async () => {
      const result = await collect(async function* () {
        // Empty
      }());
      expect(result).toEqual([]);
    });
  });
});

