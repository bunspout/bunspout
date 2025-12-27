import type { Row, Cell, XmlEvent } from '../types';

/**
 * Parses a sheet from XML events, yielding rows
 */
export async function* parseSheet(
  xmlEvents: AsyncIterable<XmlEvent>,
  getSharedString?: (index: number) => string | undefined,
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
          type: cellType === 's' || cellType === 'inlineStr' ? 'string' : undefined,
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
      } else if (event.name === 'c' && inCell && currentCell && currentRow) {
        inCell = false;
        // Add cell even if value is undefined (empty cell) - set to empty string
        if (currentCell.value === undefined) {
          currentCell.value = '';
          currentCell.type = 'string';
        }
        currentRow.cells!.push(currentCell as Cell);
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
            currentCell.type = 'number';
          } else {
            currentCell.value = text;
            if (currentCell.type === undefined) {
              currentCell.type = 'string';
            }
          }
        }
      } else if (inInlineStrText && currentCell) {
        // Inline string text element (<t> inside <is>)
        const text = event.text || '';
        currentCell.value = text;
        currentCell.type = 'string';
      }
    }
  }

  // Yield any remaining row
  if (currentRow && inRow) {
    yield currentRow as Row;
  }
}

