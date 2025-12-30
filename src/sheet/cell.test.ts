import { describe, test, expect } from 'bun:test';
import {
  cell,
  cellFromString,
  cellFromNumber,
  cellFromDate,
  cellFromBoolean,
  cellFromNull,
} from './cell';

describe('Cell Factory Functions', () => {
  describe('cell()', () => {
    test('should auto-detect string type', () => {
      const result = cell('Hello');
      expect(result.value).toBe('Hello');
      expect(result.type).toBe('string');
    });

    test('should auto-detect number type', () => {
      const result = cell(42);
      expect(result.value).toBe(42);
      expect(result.type).toBe('number');
    });

    test('should auto-detect Date type', () => {
      const date = new Date('2024-01-01');
      const result = cell(date);
      expect(result.type).toBe('date');
      expect(typeof result.value).toBe('number');
      expect(result.value).toBeGreaterThan(0);
    });

    test('should auto-detect boolean type', () => {
      const resultTrue = cell(true);
      expect(resultTrue.value).toBe(1);
      expect(resultTrue.type).toBe('boolean');

      const resultFalse = cell(false);
      expect(resultFalse.value).toBe(0);
      expect(resultFalse.type).toBe('boolean');
    });

    test('should handle null', () => {
      const result = cell(null);
      expect(result.value).toBe('');
      expect(result.type).toBeUndefined();
    });

    test('should handle undefined', () => {
      const result = cell(undefined);
      expect(result.value).toBe('');
      expect(result.type).toBeUndefined();
    });
  });

  describe('cellFromString()', () => {
    test('should create string cell', () => {
      const result = cellFromString('Test');
      expect(result.value).toBe('Test');
      expect(result.type).toBe('string');
    });
  });

  describe('cellFromNumber()', () => {
    test('should create number cell', () => {
      const result = cellFromNumber(123);
      expect(result.value).toBe(123);
      expect(result.type).toBe('number');
    });
  });

  describe('cellFromDate()', () => {
    test('should convert Date to Excel serial number', () => {
      const date = new Date('2024-01-01');
      const result = cellFromDate(date);
      expect(result.type).toBe('date');
      expect(typeof result.value).toBe('number');
      expect(result.value).toBeGreaterThan(45000); // Should be around 45323 for 2024-01-01
    });
  });

  describe('cellFromBoolean()', () => {
    test('should convert true to 1', () => {
      const result = cellFromBoolean(true);
      expect(result.value).toBe(1);
      expect(result.type).toBe('boolean');
    });

    test('should convert false to 0', () => {
      const result = cellFromBoolean(false);
      expect(result.value).toBe(0);
      expect(result.type).toBe('boolean');
    });
  });

  describe('cellFromNull()', () => {
    test('should create empty cell', () => {
      const result = cellFromNull();
      expect(result.value).toBe('');
      expect(result.type).toBeUndefined();
    });
  });
});
