/*
 * Column width utilities for Excel
 */
import type { ColumnWidthDefinition } from '@xlsx/types';
import type { Cell } from '../types';
import { resolveCell } from '../xml/cell-resolver';

/**
 * Roughly estimates Excel column width from cell content
 * Excel width is measured in "characters" (approximately 7 pixels per character for default font)
 * Formula: width ≈ (character count * 1.2) + 2 (with minimum of 2)
 */
export function estimateColumnWidth(cell: Cell): number {
  const resolved = resolveCell(cell);

  switch (resolved.t) {
    case 's': {
      // String: use character count
      const text = String(resolved.v);
      // Count characters, accounting for wide characters (rough estimate)
      const charCount = [...text].length; // counts code points properly
      // Excel formula: approximately 1.2 chars per character + padding
      return Math.max(2, Math.ceil(charCount * 1.2 + 2));
    }
    case 'n': {
      // Number: estimate based on string representation
      const text = String(resolved.v);
      return Math.max(2, Math.ceil(text.length * 1.2 + 2));
    }
    case 'd':
      // Date: typical date format is ~10-12 chars
      return 12;
    case 'b':  // TRUE / FALSE
      return 6;
    default:
      // Default for unknown types
      return 10;
  }
}

/**
 * Tracks maximum column widths while processing rows
 * All column indices are 0-based (0 = column A, 1 = column B, etc.)
 */
export class ColumnWidthTracker {
  private widths: Map<number, number> = new Map();
  private autoDetect: boolean;

  constructor(autoDetect: boolean = false) {
    this.autoDetect = autoDetect;
  }

  /**
   * Updates width for a column based on cell content
   * @param colIndex - Column index (0-based: 0 = A, 1 = B, etc.)
   * @param cell - Cell to estimate width from
   */
  updateColumnWidth(colIndex: number, cell: Cell): void {
    if (!this.autoDetect) {
      return;
    }

    const estimatedWidth = estimateColumnWidth(cell);
    const currentWidth = this.widths.get(colIndex) ?? 0;
    this.widths.set(colIndex, Math.max(currentWidth, estimatedWidth));
  }

  /**
   * Gets the tracked width for a column
   * @param colIndex - Column index (0-based: 0 = A, 1 = B, etc.)
   * @returns Width in Excel units, or undefined if not tracked
   */
  getColumnWidth(colIndex: number): number | undefined {
    return this.widths.get(colIndex);
  }

  /**
   * Gets all tracked widths
   */
  getAllWidths(): Map<number, number> {
    return new Map(this.widths);
  }
}

/**
 * Resolves final column width for a given column index
 * Priority: explicit width > auto-detected > default
 * @param colIndex - Column index (0-based: 0 = A, 1 = B, etc.)
 * @param defaultWidth - Default width to use if no specific definition matches
 * @param columnWidths - Array of column width definitions (all indices are 0-based)
 * @param trackedWidth - Auto-detected width from content (if available)
 * @returns Resolved width in Excel units, or undefined if no width should be set
 */
export function resolveColumnWidth(
  colIndex: number,
  defaultWidth: number | undefined,
  columnWidths: Array<ColumnWidthDefinition> | undefined,
  trackedWidth: number | undefined,
): number | undefined {
  let hasAutoDetectMatch = false;

  // Check for explicit column width definitions
  if (columnWidths) {
    for (const def of columnWidths) {
      // Check if this definition matches the column
      let matchesRange = false;
      if (def.columnIndex !== undefined && def.columnIndex === colIndex) {
        matchesRange = true;
      } else if (def.columnRange) {
        const { from, to } = def.columnRange;
        // Range is inclusive on both ends (0-based)
        if (colIndex >= from && colIndex <= to) {
          matchesRange = true;
        }
      }

      if (matchesRange) {
        // Early return on explicit width
        if (def.width !== undefined) {
          return def.width;
        }
        // Track if auto-detect is requested for this column
        if (def.autoDetect) {
          hasAutoDetectMatch = true;
        }
      }
    }
  }

  // Use auto-detected width if a definition requested it
  if (hasAutoDetectMatch && trackedWidth !== undefined) {
    return trackedWidth;
  }

  // Use auto-detected width if available (global auto-detect)
  if (trackedWidth !== undefined) {
    return trackedWidth;
  }

  // Fall back to default
  return defaultWidth;
}
