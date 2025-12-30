import { describe, test, expect } from 'bun:test';
import {
  columnIndexToLetter,
  columnLetterToIndex,
  getCellReference,
  parseCellReference,
} from './cell-reference';

describe('Cell Reference Utilities', () => {
  describe('columnIndexToLetter', () => {
    test('should convert single letter columns (A-Z)', () => {
      expect(columnIndexToLetter(0)).toBe('A');
      expect(columnIndexToLetter(1)).toBe('B');
      expect(columnIndexToLetter(25)).toBe('Z');
    });

    test('should convert double letter columns (AA-ZZ)', () => {
      expect(columnIndexToLetter(26)).toBe('AA');
      expect(columnIndexToLetter(27)).toBe('AB');
      expect(columnIndexToLetter(51)).toBe('AZ');
      expect(columnIndexToLetter(52)).toBe('BA');
    });

    test('should convert triple letter columns', () => {
      expect(columnIndexToLetter(702)).toBe('AAA');
      expect(columnIndexToLetter(703)).toBe('AAB');
    });
  });

  describe('columnLetterToIndex', () => {
    test('should convert single letter columns (A-Z)', () => {
      expect(columnLetterToIndex('A')).toBe(0);
      expect(columnLetterToIndex('B')).toBe(1);
      expect(columnLetterToIndex('Z')).toBe(25);
    });

    test('should convert double letter columns (AA-ZZ)', () => {
      expect(columnLetterToIndex('AA')).toBe(26);
      expect(columnLetterToIndex('AB')).toBe(27);
      expect(columnLetterToIndex('AZ')).toBe(51);
      expect(columnLetterToIndex('BA')).toBe(52);
    });

    test('should convert triple letter columns', () => {
      expect(columnLetterToIndex('AAA')).toBe(702);
      expect(columnLetterToIndex('AAB')).toBe(703);
    });

    test('should be inverse of columnIndexToLetter', () => {
      for (let i = 0; i < 1000; i++) {
        const letter = columnIndexToLetter(i);
        expect(columnLetterToIndex(letter)).toBe(i);
      }
    });
  });

  describe('getCellReference', () => {
    test('should generate correct cell references', () => {
      expect(getCellReference(1, 0)).toBe('A1');
      expect(getCellReference(1, 1)).toBe('B1');
      expect(getCellReference(2, 0)).toBe('A2');
      expect(getCellReference(26, 0)).toBe('A26');
      expect(getCellReference(1, 26)).toBe('AA1');
      expect(getCellReference(100, 51)).toBe('AZ100');
    });
  });

  describe('parseCellReference', () => {
    test('should parse single letter cell references', () => {
      expect(parseCellReference('A1')).toEqual({ rowIndex: 1, colIndex: 0 });
      expect(parseCellReference('B1')).toEqual({ rowIndex: 1, colIndex: 1 });
      expect(parseCellReference('Z1')).toEqual({ rowIndex: 1, colIndex: 25 });
    });

    test('should parse double letter cell references', () => {
      expect(parseCellReference('AA1')).toEqual({ rowIndex: 1, colIndex: 26 });
      expect(parseCellReference('AB1')).toEqual({ rowIndex: 1, colIndex: 27 });
      expect(parseCellReference('BA1')).toEqual({ rowIndex: 1, colIndex: 52 });
    });

    test('should parse various row numbers', () => {
      expect(parseCellReference('A1')).toEqual({ rowIndex: 1, colIndex: 0 });
      expect(parseCellReference('A10')).toEqual({ rowIndex: 10, colIndex: 0 });
      expect(parseCellReference('A100')).toEqual({ rowIndex: 100, colIndex: 0 });
      expect(parseCellReference('A1000')).toEqual({ rowIndex: 1000, colIndex: 0 });
    });

    test('should return null for invalid references', () => {
      expect(parseCellReference('')).toBeNull();
      expect(parseCellReference('1A')).toBeNull();
      expect(parseCellReference('A')).toBeNull();
      expect(parseCellReference('1')).toBeNull();
      expect(parseCellReference('A1B')).toBeNull();
      expect(parseCellReference('a1')).toBeNull(); // lowercase not supported
    });

    test('should be inverse of getCellReference', () => {
      const testCases = [
        { row: 1, col: 0 },
        { row: 1, col: 25 },
        { row: 1, col: 26 },
        { row: 10, col: 0 },
        { row: 100, col: 51 },
      ];

      for (const { row, col } of testCases) {
        const ref = getCellReference(row, col);
        const parsed = parseCellReference(ref);
        expect(parsed).toEqual({ rowIndex: row, colIndex: col });
      }
    });
  });
});
