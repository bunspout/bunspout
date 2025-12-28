import type { ColumnWidthDefinition, SheetColumnWidthOptions, RowHeightDefinition } from '@xlsx/types';
import { getCellReference } from '@utils/cell-reference';
import { ColumnWidthTracker } from '@utils/column-widths';
import { escapeXml } from '@utils/xml';
import type { Row, Cell } from '../types';
import { resolveCell } from './cell-resolver';

/**
 * Options for writing sheet XML
 */
export interface WriteSheetXmlOptions {
  /**
   * Function to get shared string index for a given string.
   * If provided, strings will be stored in shared strings table.
   */
  getStringIndex?: (str: string) => number;
  /**
   * Column width options for the sheet
   */
  columnWidths?: SheetColumnWidthOptions;
  /**
   * Default row height for all rows (in points)
   */
  defaultRowHeight?: number;
  /**
   * Row height definitions for specific rows/ranges
   */
  rowHeights?: RowHeightDefinition[];
}

/**
 * Options for serializing a row to XML
 */
export interface SerializeRowOptions {
  /**
   * Function to get shared string index for a given string.
   * If provided, strings will be stored in shared strings table.
   */
  getStringIndex?: (str: string) => number;
  /**
   * Column width tracker for auto-detecting column widths.
   * If provided, cell widths will be tracked during serialization.
   */
  widthTracker?: ColumnWidthTracker;
  /**
   * Resolved row height in points.
   * If provided, the row will include height attributes in the XML.
   */
  rowHeight?: number;
  /**
   * Row index to use for serialization.
   * If provided, this overrides row.rowIndex.
   * Used for auto-inferring row indices when row.rowIndex is undefined.
   */
  rowIndex?: number;
}

/**
 * Serializes a single cell to XML
 * @param cell - Cell to serialize
 * @param rowIndex - Row index (1-based: 1 = first row)
 * @param colIndex - Column index (0-based: 0 = A, 1 = B, etc.)
 * @param getStringIndex - Optional function to get shared string index
 * @returns XML string for the cell
 */
export function serializeCell(
  cell: Cell,
  rowIndex: number,
  colIndex: number,
  getStringIndex?: (str: string) => number,
): string {
  const resolved = resolveCell(cell);
  const cellRef = getCellReference(rowIndex, colIndex);
  const styleAttr = ' s="0"'; // Default style index
  const refAttr = ` r="${cellRef}"`;

  if (resolved.t === 's' && getStringIndex) {
    // Use shared strings - reference by index
    const index = getStringIndex(resolved.v as string);
    return `<c${refAttr}${styleAttr} t="s"><v>${index}</v></c>`;
  }

  if (resolved.t === 's' && !getStringIndex) {
    // Use inline strings - text stored directly in cell
    const text = escapeXml(resolved.v as string);
    return `<c${refAttr}${styleAttr} t="inlineStr"><is><t>${text}</t></is></c>`;
  }

  // Handle non-string types (number, date, boolean, etc.)
  const typeAttr = resolved.t !== 'n' ? ` t="${resolved.t}"` : '';
  const value = String(resolved.v);
  return `<c${refAttr}${styleAttr}${typeAttr}><v>${value}</v></c>`;
}

/**
 * Resolves the height for a row based on row height definitions
 * @param rowIndex - Row index (1-based)
 * @param rowHeight - Direct height on the row (if set)
 * @param rowHeights - Array of row height definitions
 * @returns Resolved height or undefined
 */
function resolveRowHeight(
  rowIndex: number,
  rowHeight: number | undefined,
  rowHeights: RowHeightDefinition[] | undefined,
): number | undefined {
  // Direct height on row takes priority
  if (rowHeight !== undefined) {
    return rowHeight;
  }

  // Check row height definitions
  if (rowHeights) {
    for (const def of rowHeights) {
      if (def.rowIndex !== undefined && def.rowIndex === rowIndex) {
        return def.height;
      } else if (def.rowRange) {
        const { from, to } = def.rowRange;
        if (rowIndex >= from && rowIndex <= to) {
          return def.height;
        }
      }
    }
  }

  return undefined;
}

/**
 * Serializes a row to XML
 */
export function serializeRow(
  row: Row,
  options?: SerializeRowOptions,
): string {
  const { getStringIndex, widthTracker, rowHeight, rowIndex: inferredRowIndex } = options ?? {};
  const rowIndex = inferredRowIndex ?? row.rowIndex ?? 1;
  const rowIndexAttr = ` r="${rowIndex}"`;

  // Calculate spans (first column to last column, 1-based)
  // Handle rows with holes (e.g., cell at col 3 but not at col 1)
  const firstColIndex = row.cells.findIndex((c) => c !== undefined && c !== null);
  const firstCol = firstColIndex >= 0 ? firstColIndex + 1 : 1;

  // Find last column with an actual cell
  let lastCol = firstCol;
  for (let i = row.cells.length - 1; i >= 0; i--) {
    if (row.cells[i] !== undefined && row.cells[i] !== null) {
      lastCol = i + 1;
      break;
    }
  }
  const spansAttr = ` spans="${firstCol}:${lastCol}"`;

  // Add height attributes if height is specified
  let heightAttrs = '';
  if (rowHeight !== undefined) {
    heightAttrs = ` ht="${rowHeight}" customHeight="1"`;
  }

  const cellsXml = row.cells
    .map((cell, colIndex) => {
      // Skip undefined/null cells (holes in the row)
      if (cell === undefined || cell === null) {
        return '';
      }
      // Track column width if auto-detection is enabled
      if (widthTracker && cell !== undefined && cell !== null) {
        widthTracker.updateColumnWidth(colIndex, cell);
      }
      return serializeCell(cell, rowIndex, colIndex, getStringIndex);
    })
    .join('');

  return `<row${rowIndexAttr}${spansAttr}${heightAttrs}>${cellsXml}</row>`;
}

/**
 * Generates the cols XML element for column width definitions
 * @param maxColumnIndex - Maximum column index (0-based: 0 = A, 1 = B, etc.)
 * @param defaultWidth - Default width for columns without specific definitions
 * @param columnWidths - Array of column width definitions (all indices are 0-based, ranges are inclusive)
 * @param trackedWidths - Map of auto-detected widths by column index (0-based)
 * @returns XML string for the cols element, or empty string if no columns need width definitions
 */
function generateColsXml(
  maxColumnIndex: number,
  defaultWidth: number | undefined,
  columnWidths: Array<ColumnWidthDefinition> | undefined,
  trackedWidths: Map<number, number>,
): string {
  if (maxColumnIndex < 0) {
    return '';
  }

  const colElements: string[] = [];
  const processedColumns = new Set<number>();

  // Process explicit column width definitions first (they take priority)
  if (columnWidths) {
    for (const def of columnWidths) {
      if (def.columnIndex !== undefined) {
        const colIndex = def.columnIndex;
        if (colIndex <= maxColumnIndex && !processedColumns.has(colIndex)) {
          const width = def.width ?? (def.autoDetect ? trackedWidths.get(colIndex) : undefined);
          if (width !== undefined) {
            colElements.push(`    <col min="${colIndex + 1}" max="${colIndex + 1}" width="${width}" customWidth="1"/>`);
            processedColumns.add(colIndex);
          }
        }
      } else if (def.columnRange) {
        const { from, to } = def.columnRange;
        const min = Math.max(0, from);
        const max = Math.min(maxColumnIndex, to);
        if (min <= max) {
          // Handle auto-detect for ranges first (each column needs individual width)
          if (def.autoDetect) {
            for (let i = min; i <= max; i++) {
              if (!processedColumns.has(i)) {
                const trackedWidth = trackedWidths.get(i);
                if (trackedWidth !== undefined) {
                  colElements.push(`    <col min="${i + 1}" max="${i + 1}" width="${trackedWidth}" customWidth="1"/>`);
                  processedColumns.add(i);
                }
              }
            }
            continue;
          }
          // Handle explicit width for ranges
          if (def.width !== undefined) {
            colElements.push(`    <col min="${min + 1}" max="${max + 1}" width="${def.width}" customWidth="1"/>`);
            for (let i = min; i <= max; i++) {
              processedColumns.add(i);
            }
          }
        }
      }
    }
  }

  // Process auto-detected widths for columns not explicitly defined
  for (let colIndex = 0; colIndex <= maxColumnIndex; colIndex++) {
    if (!processedColumns.has(colIndex)) {
      const trackedWidth = trackedWidths.get(colIndex);
      if (trackedWidth !== undefined) {
        colElements.push(`    <col min="${colIndex + 1}" max="${colIndex + 1}" width="${trackedWidth}" customWidth="1"/>`);
        processedColumns.add(colIndex);
      } else if (defaultWidth !== undefined) {
        colElements.push(`    <col min="${colIndex + 1}" max="${colIndex + 1}" width="${defaultWidth}" customWidth="1"/>`);
        processedColumns.add(colIndex);
      }
    }
  }

  // If we have a default width but no explicit columns, create a single col element for all columns
  if (colElements.length === 0 && defaultWidth !== undefined) {
    colElements.push(`    <col min="1" max="${maxColumnIndex + 1}" width="${defaultWidth}" customWidth="1"/>`);
  }

  if (colElements.length === 0) {
    return '';
  }

  return `<cols>\n${colElements.join('\n')}\n  </cols>`;
}

/**
 * Generates sheetFormatPr XML for default column width and row height
 */
function generateSheetFormatPr(
  defaultWidth: number | undefined,
  defaultRowHeight: number | undefined,
): string {
  const attrs: string[] = [];
  if (defaultWidth !== undefined) {
    attrs.push(`defaultColWidth="${defaultWidth}"`);
  }
  if (defaultRowHeight !== undefined) {
    attrs.push(`defaultRowHeight="${defaultRowHeight}"`);
  }
  if (attrs.length === 0) {
    return '';
  }
  return `    <sheetFormatPr ${attrs.join(' ')}/>`;
}

/**
 * Writes sheet XML from an async iterable of rows
 *
 * Performance optimization:
 * - Fast path: When column widths are NOT needed, rows are streamed directly without buffering.
 *   This allows processing of very large sheets with constant memory usage.
 * - Slow path: When column widths ARE needed, rows must be buffered to calculate maxColumnIndex
 *   and generate the cols XML before sheetData. This is unavoidable because:
 *   1. We need maxColumnIndex to generate cols XML with correct column ranges
 *   2. We need actual row data for auto-detection width tracking
 *   3. AsyncIterables from streams cannot be re-iterated
 *
 * For re-readable sources (e.g., generator functions), callers can implement a two-pass
 * approach themselves to avoid buffering in this function.
 */
export async function* writeSheetXml(
  rows: AsyncIterable<Row>,
  options?: WriteSheetXmlOptions,
): AsyncIterable<string> {
  yield '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

  const getStringIndex = options?.getStringIndex;
  const columnWidthOptions = options?.columnWidths;
  const defaultRowHeight = options?.defaultRowHeight;
  const rowHeights = options?.rowHeights;

  // Check if we need per-column width definitions (cols XML)
  // A global defaultColumnWidth only needs sheetFormatPr, not cols XML
  const needsColsXml =
    (columnWidthOptions?.columnWidths !== undefined && columnWidthOptions.columnWidths.length > 0) ||
    columnWidthOptions?.autoDetectColumnWidth === true;

  // Initialize column width tracker if auto-detection is enabled (globally or per-column)
  const autoDetect = columnWidthOptions?.autoDetectColumnWidth ?? false;
  const hasPerColumnAutoDetect = columnWidthOptions?.columnWidths?.some(def => def.autoDetect) ?? false;
  const needsWidthTracking = autoDetect || hasPerColumnAutoDetect;
  const widthTracker = new ColumnWidthTracker(needsWidthTracking);

  // Fast path: no per-column overrides needed - stream directly
  // This covers both: no widths at all, and defaultColumnWidth only (which just needs sheetFormatPr)
  if (!needsColsXml) {
    // Yield sheetFormatPr if we have a default width or row height
    if (columnWidthOptions?.defaultColumnWidth !== undefined || defaultRowHeight !== undefined) {
      yield generateSheetFormatPr(columnWidthOptions?.defaultColumnWidth, defaultRowHeight);
    }
    yield '<sheetData>';

    let currentRowNumber = 1;
    for await (const row of rows) {
      // Auto-assign row index if not provided, based on order
      const rowIndex = row.rowIndex ?? currentRowNumber;
      const resolvedHeight = resolveRowHeight(rowIndex, row.height, rowHeights);
      yield serializeRow(row, {
        getStringIndex,
        widthTracker,
        rowHeight: resolvedHeight,
        rowIndex,
      });
      // Increment for next row (only if rowIndex wasn't explicitly set)
      if (row.rowIndex === undefined) {
        currentRowNumber++;
      } else {
        // If explicit rowIndex is set, use it + 1 as next default (allows for gaps)
        currentRowNumber = row.rowIndex + 1;
      }
    }

    yield '</sheetData></worksheet>';
    return;
  }

  // Slow path: need to buffer rows to calculate maxColumnIndex and generate cols XML
  // This is unavoidable when column widths are needed because:
  // 1. We need maxColumnIndex to generate cols XML with correct column ranges
  // 2. We need actual row data for auto-detection width tracking
  // 3. AsyncIterables from streams cannot be re-iterated
  const allRows: Row[] = [];
  let maxColumnIndex = -1;

  for await (const row of rows) {
    allRows.push(row);
    maxColumnIndex = Math.max(maxColumnIndex, row.cells.length - 1);
    // Track widths if auto-detection is enabled
    if (needsWidthTracking) {
      row.cells.forEach((cell, colIndex) => {
        if (cell !== undefined && cell !== null) {
          widthTracker.updateColumnWidth(colIndex, cell);
        }
      });
    }
  }

  // Generate column width XML if needed
  const sheetFormatPr = generateSheetFormatPr(columnWidthOptions?.defaultColumnWidth, defaultRowHeight);
  // Only generate cols XML if we have per-column definitions or auto-detection
  // A global defaultColumnWidth alone doesn't need cols XML (only sheetFormatPr)
  const colsXml = needsColsXml
    ? generateColsXml(
      maxColumnIndex,
      columnWidthOptions?.defaultColumnWidth,
      columnWidthOptions?.columnWidths,
      widthTracker.getAllWidths(),
    )
    : '';

  // Yield sheetFormatPr and cols before sheetData
  if (sheetFormatPr) {
    yield sheetFormatPr;
  }
  if (colsXml) {
    yield colsXml;
  }

  yield '<sheetData>';

  // Yield all rows with auto-assigned row indices
  let currentRowNumber = 1;
  for (const row of allRows) {
    // Auto-assign row index if not provided, based on order
    const rowIndex = row.rowIndex ?? currentRowNumber;
    const resolvedHeight = resolveRowHeight(rowIndex, row.height, rowHeights);
    yield serializeRow(row, {
      getStringIndex,
      widthTracker,
      rowHeight: resolvedHeight,
      rowIndex,
    });
    // Increment for next row (only if rowIndex wasn't explicitly set)
    if (row.rowIndex === undefined) {
      currentRowNumber++;
    } else {
      // If explicit rowIndex is set, use it + 1 as next default (allows for gaps)
      currentRowNumber = row.rowIndex + 1;
    }
  }

  yield '</sheetData></worksheet>';
}

