import type { ReadOptions } from '@xlsx/types';
import { convertExcelTimestamp } from '@utils/dates';
import type { Cell, Row, XmlEvent } from '../types';

/**
 * Parses a sheet from XML events, yielding rows
 */
export async function* parseSheet(
  xmlEvents: AsyncIterable<XmlEvent>,
  getSharedString?: (index: number) => string | undefined,
  options?: ReadOptions,
): AsyncIterable<Row> {
  let currentRow: Partial<Row> | null = null;
  let currentCell: Partial<Cell> | null = null;
  let inRow = false;
  let inCell = false;
  let inValue = false;
  let inInlineStr = false;
  let inInlineStrText = false;

  for await (const event of xmlEvents) {
    if (event.type === 'startElement') {
      if (event.name === 'row') {
        inRow = true;
        currentRow = {
          cells: [],
          rowIndex: event.attributes?.r ? parseInt(event.attributes.r, 10) : undefined,
        };
      } else if (event.name === 'c' && inRow) {
        inCell = true;
        const cellType = event.attributes?.t;
        currentCell = {
          type: cellType === 's' || cellType === 'inlineStr' ? 'string' :
            cellType === 'd' ? 'date' :
              cellType === 'b' ? 'boolean' : undefined,
        };
        // Reset inline string state when starting a new cell
        inInlineStr = false;
        inInlineStrText = false;
      } else if (event.name === 'v' && inCell && !inInlineStr) {
        inValue = true;
      } else if (event.name === 'is' && inCell) {
        inInlineStr = true;
      } else if (event.name === 't' && inInlineStr) {
        inInlineStrText = true;
      }
    } else if (event.type === 'endElement') {
      if (event.name === 'row' && inRow && currentRow) {
        inRow = false;
        yield currentRow as Row;
        currentRow = null;
      } else if (event.name === 'c' && inCell && currentCell && currentRow && currentRow.cells) {
        inCell = false;
        // Add cell even if value is undefined (empty cell) - set to empty string
        if (currentCell.value === undefined) {
          currentCell.value = '';
          currentCell.type = 'string';
        }

        // Convert date cells if options are provided
        if (options && currentCell.type === 'date' && typeof currentCell.value === 'number') {
          try {
            currentCell.value = convertExcelTimestamp(currentCell.value, options.use1904Dates ?? false);
          } catch (error) {
            // If conversion fails, keep the original numeric value
            console.warn(`Failed to convert Excel date ${currentCell.value}:`, error);
          }
        }

        // Convert boolean cells
        if (currentCell.type === 'boolean' && typeof currentCell.value === 'number') {
          currentCell.value = currentCell.value === 1;
        }

        const cell: Cell = {
          value: currentCell.value ?? '',
          ...(currentCell.type !== undefined && { type: currentCell.type }),
        };
        currentRow.cells.push(cell);
        currentCell = null;
        inInlineStr = false;
        inInlineStrText = false;
      } else if (event.name === 'v' && inValue) {
        inValue = false;
      } else if (event.name === 't' && inInlineStrText) {
        inInlineStrText = false;
      } else if (event.name === 'is' && inInlineStr) {
        inInlineStr = false;
      }
    } else if (event.type === 'text') {
      if (inValue && currentCell && !inInlineStr) {
        // Regular value element (<v>)
        const text = event.text || '';

        // If cell type is 's' (shared string), look up the string by index
        if (currentCell.type === 'string' && getSharedString) {
          const index = parseInt(text, 10);
          if (!isNaN(index)) {
            const sharedString = getSharedString(index);
            if (sharedString !== undefined) {
              currentCell.value = sharedString;
            } else {
              currentCell.value = text; // Fallback if index not found
            }
          } else {
            currentCell.value = text;
          }
        } else {
          // Try to parse as number if not explicitly a string
          if (currentCell.type !== 'string' && !isNaN(Number(text)) && text.trim() !== '') {
            currentCell.value = Number(text);
            // Only set type to 'number' if it wasn't explicitly set from XML attributes
            if (currentCell.type === undefined) {
              currentCell.type = 'number';
            }
          } else {
            currentCell.value = text;
            if (currentCell.type === undefined) {
              currentCell.type = 'string';
            }
          }
        }
      } else if (inInlineStrText && currentCell) {
        // Inline string text element (<t> inside <is>)
        currentCell.value = event.text || '';
        currentCell.type = 'string';
      }
    }
  }

  // Yield any remaining row
  if (currentRow && inRow) {
    yield currentRow as Row;
  }
}

