/*
 * Date conversion utilities for Excel date handling
 */

/**
 * Converts an Excel timestamp (serial number) to a JavaScript Date object.
 * Excel stores dates as the number of days since a base date, with fractional
 * parts representing time of day.
 *
 * @param excelValue - Excel serial number (days since base date)
 * @param use1904Dates - Whether to use 1904-based calendar (default: false for 1900-based)
 * @returns JavaScript Date object
 */
export function convertExcelTimestamp(
  excelValue: number,
  use1904Dates: boolean = false,
): Date {
  // Validate input range
  const minValue = use1904Dates ? -695055 : -693593;
  const maxValue = use1904Dates ? 2957003.9999884 : 2958465.9999884;

  if (excelValue < minValue || excelValue > maxValue) {
    throw new Error(
      `Excel timestamp ${excelValue} is outside valid range [${minValue}, ${maxValue}]`,
    );
  }

  // Split into integer days and fractional time
  const days = Math.floor(excelValue);
  const timeFraction = excelValue - days;

  if (use1904Dates) {
    // 1904-based calendar: Excel day 0 = January 1, 1904
    const baseDate = new Date(1904, 0, 1); // January 1, 1904
    const millisecondsInDay = 24 * 60 * 60 * 1000;
    return new Date(baseDate.getTime() + days * millisecondsInDay + timeFraction * millisecondsInDay);
  } else {
    // 1900-based calendar with Excel's leap year bug
    // Excel treats 1900 as a leap year (it wasn't), so Feb 29, 1900 exists in Excel
    // This means dates from March 1, 1900 onward are offset by +1 day

    let actualDays = days;
    if (days >= 60) {
      // For dates on or after March 1, 1900 (Excel day 61), subtract 1 to compensate
      // Excel thinks Feb 29, 1900 exists (day 60), but it doesn't in reality
      actualDays = days - 1;
    }

    // Base is December 31, 1899 (Excel day 0)
    // Excel day 1 = January 1, 1900 = December 31 + 1 day
    const baseDate = new Date(1899, 11, 31); // December 31, 1899
    const millisecondsInDay = 24 * 60 * 60 * 1000;

    return new Date(baseDate.getTime() + actualDays * millisecondsInDay + timeFraction * millisecondsInDay);
  }
}

/**
 * Checks if a number represents a valid Excel date timestamp.
 * This is a heuristic check and may have false positives.
 *
 * @param value - Numeric value to check
 * @param use1904Dates - Whether to use 1904-based calendar
 * @returns true if the value is likely an Excel date
 */
export function isValidExcelDate(value: number, use1904Dates: boolean = false): boolean {
  if (isNaN(value) || !isFinite(value)) {
    return false;
  }
  try {
    convertExcelTimestamp(value, use1904Dates);
    return true;
  } catch {
    return false;
  }
}
