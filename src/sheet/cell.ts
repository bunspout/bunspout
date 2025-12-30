import type { Cell } from 'types';

/**
 * Converts a Date to Excel serial number
 * Excel serial date: days since December 31, 1899 (day 0)
 * Day 1 = January 1, 1900
 */
function dateToExcelSerial(date: Date): number {
  const epoch = new Date(1899, 11, 31); // December 31, 1899 (Excel day 0)
  const diff = date.getTime() - epoch.getTime();
  const days = diff / (1000 * 60 * 60 * 24);
  // Excel incorrectly treats 1900 as a leap year (Feb 29, 1900 exists in Excel)
  // For dates on or after March 1, 1900, add 1 to account for the phantom Feb 29
  if (date >= new Date(1900, 2, 1)) {
    return days + 1;
  }
  return days;
}

/**
 * Auto-detects type from value and creates a Cell
 */
export function cell(
  value: string | number | Date | boolean | null | undefined,
): Cell {
  if (value === null || value === undefined) {
    return { value: '', type: undefined };
  }
  if (typeof value === 'string') {
    return { value, type: 'string' };
  }
  if (typeof value === 'number') {
    return { value, type: 'number' };
  }
  if (value instanceof Date) {
    return { value: dateToExcelSerial(value), type: 'date' };
  }
  if (typeof value === 'boolean') {
    return { value: value ? 1 : 0, type: 'boolean' };
  }
  return { value: String(value), type: 'string' };
}

/**
 * Creates a string cell
 */
export function cellFromString(value: string): Cell {
  return { value, type: 'string' };
}

/**
 * Creates a number cell
 */
export function cellFromNumber(value: number): Cell {
  return { value, type: 'number' };
}

/**
 * Creates a date cell (converts Date to Excel serial number)
 */
export function cellFromDate(date: Date): Cell {
  return { value: dateToExcelSerial(date), type: 'date' };
}

/**
 * Creates a boolean cell (converts to 1/0)
 */
export function cellFromBoolean(value: boolean): Cell {
  return { value: value ? 1 : 0, type: 'boolean' };
}

/**
 * Creates an empty/null cell
 */
export function cellFromNull(): Cell {
  return { value: '', type: undefined };
}
