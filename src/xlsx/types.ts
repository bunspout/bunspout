import type { Row } from '../types';

export interface WorkbookProperties {
  title?: string | null;
  subject?: string | null;
  application?: string | null;
  creator?: string | null;
  lastModifiedBy?: string | null;
  keywords?: string | null;
  description?: string | null;
  category?: string | null;
  language?: string | null;
  customProperties?: Record<string, string>;
}

/**
 * Column width definition for a sheet
 * All column indices are 0-based (0 = column A, 1 = column B, etc.)
 */
export interface ColumnWidthDefinition {
  /**
   * Single column index (0-based).
   * Mutually exclusive with columnRange.
   * @example 0 for column A, 1 for column B
   */
  columnIndex?: number;
  /**
   * Range of columns (0-based, inclusive on both ends).
   * Mutually exclusive with columnIndex.
   * @example { from: 0, to: 2 } covers columns A, B, and C
   */
  columnRange?: { from: number; to: number };
  /**
   * Explicit width in Excel units (characters).
   * Excel default is approximately 8.43 characters.
   */
  width?: number;
  /**
   * Auto-detect width from cell content.
   * When true, estimates width based on the longest content in the column.
   */
  autoDetect?: boolean;
}

export interface SheetColumnWidthOptions {
  defaultColumnWidth?: number; // Default width for all columns
  columnWidths?: ColumnWidthDefinition[]; // Override specific columns/ranges
  autoDetectColumnWidth?: boolean; // Auto-detect widths for all columns
}

/**
 * Row height definition for a sheet
 * All row indices are 1-based
 */
export interface RowHeightDefinition {
  /**
   * Single row index (1-based).
   * Mutually exclusive with rowRange.
   */
  rowIndex?: number;
  /**
   * Range of rows (1-based, inclusive on both ends).
   * Mutually exclusive with rowIndex.
   * @example { from: 1, to: 3 } covers rows 1, 2, and 3
   */
  rowRange?: { from: number; to: number };
  /**
   * Explicit height in points.
   * Excel default is approximately 15 points.
   */
  height?: number;
}

export interface WorkbookDefinition {
  sheets: SheetDefinition[];
  properties?: WorkbookProperties;
}

export interface SheetDefinition {
  name: string;
  rows: AsyncIterable<Row>;
  hidden?: boolean; // If true, sheet will be hidden in Excel
  defaultColumnWidth?: number; // Default width for all columns
  columnWidths?: ColumnWidthDefinition[]; // Override specific columns/ranges
  autoDetectColumnWidth?: boolean; // Auto-detect widths for all columns
  defaultRowHeight?: number; // Default height for all rows (in points)
  rowHeights?: RowHeightDefinition[]; // Override specific rows/ranges
}

export interface WriterOptions {
  sharedStrings?: 'inline' | 'shared'; // Default: 'inline'
}

export interface ReadOptions {
  /**
   * Use 1904-based calendar instead of 1900-based calendar for date parsing.
   * Excel files can use either calendar system.
   * @default false
   */
  use1904Dates?: boolean;
  /**
   * Skip empty rows when reading sheets.
   * When true (default), rows with no cells or all empty cells are not yielded.
   * When false, empty rows are returned as rows with empty string cells.
   * @default true
   */
  skipEmptyRows?: boolean;
  /**
   * Format dates as strings according to their Excel format codes.
   * When false (default), dates are returned as Date objects.
   * When true, dates are returned as formatted strings matching Excel's display format.
   * @default false
   */
  shouldFormatDates?: boolean;
}
