import type * as yauzl from 'yauzl';
import { readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import type { ColumnWidthDefinition, RowHeightDefinition } from './types';

export interface SheetProperties {
  hidden?: boolean;
  defaultColumnWidth?: number;
  columnWidths?: ColumnWidthDefinition[];
  defaultRowHeight?: number;
  rowHeights?: RowHeightDefinition[];
}

/**
 * Parses worksheet XML to extract sheet properties (column widths, row heights, etc.)
 */
export async function parseSheetProperties(
  zipEntry: ZipEntry,
  zipFile: yauzl.ZipFile,
): Promise<SheetProperties> {
  const properties: SheetProperties = {};
  const columnWidths: ColumnWidthDefinition[] = [];
  const rowHeights: RowHeightDefinition[] = [];

  let inCols = false;
  let inCol = false;
  let inSheetData = false;
  let inRow = false;

  let currentColMin: number | null = null;
  let currentColMax: number | null = null;
  let currentColWidth: number | null = null;

  let currentRowIndex: number | null = null;
  let currentRowHeight: number | null = null;

  for await (const event of parseXmlEvents(readZipEntry(zipEntry, zipFile))) {
    if (event.type === 'startElement') {
      if (event.name === 'sheetFormatPr') {
        const defaultColWidth = event.attributes?.defaultColWidth;
        const defaultRowHeight = event.attributes?.defaultRowHeight;
        if (defaultColWidth) {
          properties.defaultColumnWidth = parseFloat(defaultColWidth);
        }
        if (defaultRowHeight) {
          properties.defaultRowHeight = parseFloat(defaultRowHeight);
        }
      } else if (event.name === 'cols') {
        inCols = true;
      } else if (event.name === 'col' && inCols) {
        inCol = true;
        const min = event.attributes?.min;
        const max = event.attributes?.max;
        const width = event.attributes?.width;
        if (min) currentColMin = parseInt(min, 10) - 1; // Convert to 0-based
        if (max) currentColMax = parseInt(max, 10) - 1; // Convert to 0-based
        if (width) currentColWidth = parseFloat(width);
      } else if (event.name === 'sheetData') {
        inSheetData = true;
      } else if (event.name === 'row' && inSheetData) {
        inRow = true;
        const r = event.attributes?.r;
        const ht = event.attributes?.ht;
        if (r) currentRowIndex = parseInt(r, 10);
        if (ht) currentRowHeight = parseFloat(ht);
      }
    } else if (event.type === 'endElement') {
      if (event.name === 'cols') {
        inCols = false;
      } else if (event.name === 'col' && inCol) {
        inCol = false;
        if (currentColMin !== null && currentColMax !== null && currentColWidth !== null) {
          if (currentColMin === currentColMax) {
            // Single column
            columnWidths.push({
              columnIndex: currentColMin,
              width: currentColWidth,
            });
          } else {
            // Column range
            columnWidths.push({
              columnRange: { from: currentColMin, to: currentColMax },
              width: currentColWidth,
            });
          }
        }
        currentColMin = null;
        currentColMax = null;
        currentColWidth = null;
      } else if (event.name === 'row' && inRow) {
        inRow = false;
        if (currentRowIndex !== null && currentRowHeight !== null) {
          rowHeights.push({
            rowIndex: currentRowIndex,
            height: currentRowHeight,
          });
        }
        currentRowIndex = null;
        currentRowHeight = null;
      }
    }
  }

  if (columnWidths.length > 0) {
    properties.columnWidths = columnWidths;
  }
  if (rowHeights.length > 0) {
    properties.rowHeights = rowHeights;
  }

  return properties;
}
