import { describe, test, expect } from 'bun:test';
import { convertExcelTimestamp, isValidExcelDate } from './dates';

describe('Date Conversion Utilities', () => {
  describe('convertExcelTimestamp', () => {
    test('should convert 1900-based Excel dates correctly', () => {
      // Excel date 1 = January 1, 1900
      const date1 = convertExcelTimestamp(1, false);
      expect(date1.getFullYear()).toBe(1900);
      expect(date1.getMonth()).toBe(0); // January (0-based)
      expect(date1.getDate()).toBe(1);

      // Excel date 2 = January 2, 1900
      const date2 = convertExcelTimestamp(2, false);
      expect(date2.getFullYear()).toBe(1900);
      expect(date2.getMonth()).toBe(0);
      expect(date2.getDate()).toBe(2);

      // Test date with time: 1.5 = January 1, 1900 at noon
      const dateWithTime = convertExcelTimestamp(1.5, false);
      expect(dateWithTime.getFullYear()).toBe(1900);
      expect(dateWithTime.getMonth()).toBe(0); // January (0-based)
      expect(dateWithTime.getDate()).toBe(1);
      expect(dateWithTime.getHours()).toBe(12);
    });

    test('should handle Excel leap year bug correctly', () => {
      // Excel incorrectly treats 1900 as a leap year
      // February 29, 1900 should not exist but Excel thinks it does
      // Date 60 = March 1, 1900 (adjusted for leap year bug)
      const march1 = convertExcelTimestamp(61, false);
      expect(march1.getFullYear()).toBe(1900);
      expect(march1.getMonth()).toBe(2); // March (0-based)
      expect(march1.getDate()).toBe(1);
    });

    test('should convert 1904-based Excel dates correctly', () => {
      // Excel date 1 with 1904 calendar = January 2, 1904
      const date1_1904 = convertExcelTimestamp(1, true);
      expect(date1_1904.getFullYear()).toBe(1904);
      expect(date1_1904.getMonth()).toBe(0); // January
      expect(date1_1904.getDate()).toBe(2);

      // Excel date 0 with 1904 calendar = January 1, 1904
      const date0_1904 = convertExcelTimestamp(0, true);
      expect(date0_1904.getFullYear()).toBe(1904);
      expect(date0_1904.getMonth()).toBe(0);
      expect(date0_1904.getDate()).toBe(1);
    });

    test('should handle fractional time correctly', () => {
      // 0.5 = noon
      const noon = convertExcelTimestamp(44239.5, false);
      expect(noon.getHours()).toBe(12);
      expect(noon.getMinutes()).toBe(0);

      // 0.25 = 6 AM
      const morning = convertExcelTimestamp(44239.25, false);
      expect(morning.getHours()).toBe(6);
      expect(morning.getMinutes()).toBe(0);
    });

    test('should throw error for invalid timestamps', () => {
      expect(() => convertExcelTimestamp(-700000, false)).toThrow();
      expect(() => convertExcelTimestamp(3000000, false)).toThrow();
    });
  });

  describe('isValidExcelDate', () => {
    test('should return true for valid dates', () => {
      expect(isValidExcelDate(1, false)).toBe(true);
      expect(isValidExcelDate(44239, false)).toBe(true);
      expect(isValidExcelDate(1, true)).toBe(true);
    });

    test('should return false for invalid dates', () => {
      expect(isValidExcelDate(-700000, false)).toBe(false);
      expect(isValidExcelDate(3000000, false)).toBe(false);
      expect(isValidExcelDate(NaN, false)).toBe(false);
    });
  });
});
