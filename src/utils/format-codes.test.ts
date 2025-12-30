import { describe, test, expect } from 'bun:test';
import {
  convertExcelFormatToDateFns,
  getBuiltInFormatCode,
  isBuiltInDateFormat,
  isDateFormatCode,
} from './format-codes';

describe('Format Code Utilities', () => {
  describe('isBuiltInDateFormat', () => {
    test('should identify built-in date format IDs', () => {
      expect(isBuiltInDateFormat(14)).toBe(true); // mm/dd/yyyy
      expect(isBuiltInDateFormat(15)).toBe(true); // d-mmm-yy
      expect(isBuiltInDateFormat(22)).toBe(true); // m/d/yy h:mm
      expect(isBuiltInDateFormat(13)).toBe(false); // Not a date format
      expect(isBuiltInDateFormat(0)).toBe(false); // General format
    });
  });

  describe('isDateFormatCode', () => {
    test('should identify date format codes', () => {
      expect(isDateFormatCode('MM/DD/YYYY')).toBe(true);
      expect(isDateFormatCode('dd-mmm-yy')).toBe(true);
      expect(isDateFormatCode('h:mm AM/PM')).toBe(true);
      expect(isDateFormatCode('yyyy-mm-dd')).toBe(true);
      expect(isDateFormatCode('General')).toBe(false);
      expect(isDateFormatCode('0.00')).toBe(false);
    });

    test('should handle format codes with locale prefixes', () => {
      expect(isDateFormatCode('[$-409]MM/DD/YYYY')).toBe(true);
      expect(isDateFormatCode('[$-409]h:mm AM/PM')).toBe(true);
    });

    test('should not treat duration formats as dates', () => {
      expect(isDateFormatCode('[h]:mm:ss')).toBe(false);
      expect(isDateFormatCode('[mm]:ss')).toBe(false);
    });

    test('should handle empty or null format codes', () => {
      expect(isDateFormatCode('')).toBe(false);
    });
  });

  describe('convertExcelFormatToDateFns', () => {
    test('should convert basic date formats', () => {
      expect(convertExcelFormatToDateFns('MM/DD/YYYY')).toBe('MM/dd/yyyy');
      expect(convertExcelFormatToDateFns('dd-mmm-yy')).toBe('dd-MMM-yy');
      expect(convertExcelFormatToDateFns('yyyy-mm-dd')).toBe('yyyy-MM-dd');
    });

    test('should convert time formats', () => {
      expect(convertExcelFormatToDateFns('h:mm AM/PM')).toBe('h:mm a');
      expect(convertExcelFormatToDateFns('hh:mm:ss AM/PM')).toBe('hh:mm:ss a');
      expect(convertExcelFormatToDateFns('h:mm')).toBe('H:mm');
      expect(convertExcelFormatToDateFns('hh:mm:ss')).toBe('HH:mm:ss');
    });

    test('should convert month formats', () => {
      expect(convertExcelFormatToDateFns('mmmm')).toBe('MMMM');
      expect(convertExcelFormatToDateFns('mmm')).toBe('MMM');
      expect(convertExcelFormatToDateFns('mm')).toBe('MM');
      expect(convertExcelFormatToDateFns('m')).toBe('M');
    });

    test('should convert day formats', () => {
      expect(convertExcelFormatToDateFns('dddd')).toBe('EEEE');
      expect(convertExcelFormatToDateFns('ddd')).toBe('EEE');
      expect(convertExcelFormatToDateFns('dd')).toBe('dd');
      expect(convertExcelFormatToDateFns('d')).toBe('d');
    });

    test('should handle locale prefixes', () => {
      expect(convertExcelFormatToDateFns('[$-409]MM/DD/YYYY')).toBe('MM/dd/yyyy');
      expect(convertExcelFormatToDateFns('[$-409]h:mm AM/PM')).toBe('h:mm a');
    });

    test('should handle format sections separated by semicolon', () => {
      expect(convertExcelFormatToDateFns('MM/DD/YYYY;@')).toBe('MM/dd/yyyy');
      expect(convertExcelFormatToDateFns('h:mm AM/PM;General')).toBe('h:mm a');
    });

    test('should preserve quoted text', () => {
      expect(convertExcelFormatToDateFns('"Date: "MM/DD/YYYY')).toContain('Date: ');
    });

    test('should use default format for empty input', () => {
      expect(convertExcelFormatToDateFns('')).toBe('yyyy-MM-dd');
    });

    test('should handle complex formats', () => {
      expect(convertExcelFormatToDateFns('dddd mmmm dd, yy')).toBe('EEEE MMMM dd, yy');
      expect(convertExcelFormatToDateFns('m/d/yy h:mm')).toBe('M/d/yy H:mm');
    });

    test('should explicitly reject elapsed time formats', () => {
      // Elapsed time formats like [h], [m], [s] represent durations, not clock times
      // They should return a standard format rather than being mis-converted
      expect(convertExcelFormatToDateFns('[h]:mm:ss')).toBe('HH:mm:ss');
      expect(convertExcelFormatToDateFns('[mm]:ss')).toBe('HH:mm:ss');
      expect(convertExcelFormatToDateFns('[h]')).toBe('HH:mm:ss');
    });
  });

  describe('getBuiltInFormatCode', () => {
    test('should return format codes for built-in date formats', () => {
      expect(getBuiltInFormatCode(14)).toBe('mm/dd/yyyy');
      expect(getBuiltInFormatCode(15)).toBe('d-mmm-yy');
      expect(getBuiltInFormatCode(18)).toBe('h:mm AM/PM');
      expect(getBuiltInFormatCode(22)).toBe('m/d/yy h:mm');
    });

    test('should return null for non-date format IDs', () => {
      expect(getBuiltInFormatCode(0)).toBeNull();
      expect(getBuiltInFormatCode(1)).toBeNull();
      expect(getBuiltInFormatCode(13)).toBeNull();
    });
  });
});
