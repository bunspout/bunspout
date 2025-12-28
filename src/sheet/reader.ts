import type { ReadOptions } from '@xlsx/types';
import { parseCellReference } from '@utils/cell-reference';
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
  let inRichTextRun = false; // Track <r> elements (rich text runs)
  let inRubyPhonetic = false; // Skip pronunciation data
  let expectedColumnCount: number | null = null; // From spans attribute (1-based last column)
  let currentCellColIndex: number | undefined = undefined; // Column index from cell r attribute
  let explicitlySetColumns: Set<number> | null = null; // Track which column indices have been explicitly set via r attribute
  let inlineStringBuffer: string = ''; // Accumulate text from multiple rich text runs in inline strings

  for await (const event of xmlEvents) {
    if (event.type === 'startElement') {
      if (event.name === 'row') {
        inRow = true;
        // Parse spans attribute to determine expected column count
        // Format: spans="firstCol:lastCol" (1-based)
        const spans = event.attributes?.spans;
        if (spans) {
          const match = spans.match(/^(\d+):(\d+)$/);
          if (match && match.length > 2 && match[2]) {
            const lastCol = parseInt(match[2], 10);
            expectedColumnCount = lastCol; // Store 1-based last column
          }
        } else {
          expectedColumnCount = null;
        }
        currentRow = {
          cells: [],
          rowIndex: event.attributes?.r ? parseInt(event.attributes.r, 10) : undefined,
        };
        explicitlySetColumns = new Set<number>(); // Reset tracking for new row
      } else if (event.name === 'c' && inRow) {
        inCell = true;
        const cellType = event.attributes?.t;
        const cellRef = event.attributes?.r; // Cell reference like "A1", "B1", etc.
        currentCellColIndex = cellRef ? parseCellReference(cellRef)?.colIndex : undefined;
        currentCell = {
          type: cellType === 's' || cellType === 'inlineStr' ? 'string' :
            cellType === 'd' ? 'date' :
              cellType === 'b' ? 'boolean' : undefined,
        };
        // Reset inline string state when starting a new cell
        inInlineStr = false;
        inInlineStrText = false;
        inRichTextRun = false;
        inRubyPhonetic = false;
        inlineStringBuffer = ''; // Reset text accumulation buffer
      } else if (event.name === 'v' && inCell && !inInlineStr) {
        inValue = true;
      } else if (event.name === 'is' && inCell) {
        inInlineStr = true;
        inlineStringBuffer = ''; // Initialize text accumulation buffer for inline string
      } else if (event.name === 'r' && inInlineStr) {
        inRichTextRun = true;
      } else if (event.name === 't' && !inRubyPhonetic) {
        // Only extract <t> elements whose immediate parent is allowed
        // Reference: "only consider the nodes whose parents are '<si>' or '<r>'"
        // For inline strings, allow <t> directly under <is> as well for compatibility
        if (inRichTextRun) {
          inInlineStrText = true;
        } else if (inInlineStr) {
          // Allow <t> directly under <is> for inline strings (extension of reference logic)
          inInlineStrText = true;
        }
        // Skip <t> elements under <rPh> or other containers
      } else if (event.name === 'rPh' && (inInlineStr || inRichTextRun)) {
        inRubyPhonetic = true; // Skip pronunciation data
      }
    } else if (event.type === 'endElement') {
      if (event.name === 'row' && inRow && currentRow) {
        // If spans attribute was present and we have cells, pad cells array to match expected width
        // Only pad if there are actual cells (empty rows shouldn't be padded)
        if (expectedColumnCount !== null && currentRow.cells && currentRow.cells.length > 0) {
          // expectedColumnCount is 1-based, so we need that many cells (indices 0 to expectedColumnCount-1)
          while (currentRow.cells.length < expectedColumnCount) {
            currentRow.cells.push({
              value: '',
              type: 'string',
            });
          }
        }
        inRow = false;
        yield currentRow as Row;
        currentRow = null;
        expectedColumnCount = null;
        explicitlySetColumns = null;
      } else if (event.name === 'c' && inCell && currentCell && currentRow && currentRow.cells) {
        inCell = false;
        // If we were processing an inline string, use accumulated text (even if empty)
        if (inInlineStr && currentCell.value === undefined) {
          currentCell.value = inlineStringBuffer;
          currentCell.type = 'string';
        }
        // Add cell even if value is undefined (empty cell) - set to empty string
        if (currentCell.value === undefined) {
          currentCell.value = '';
          currentCell.type = 'string';
        }

        // Convert date cells if options are provided
        if (options && currentCell.type === 'date' && typeof currentCell.value === 'number') {
          try {
            currentCell.value = convertExcelTimestamp(currentCell.value, options.use1904Dates ?? false);
          } catch {
            // If conversion fails, return null for invalid dates
            currentCell.value = null;
          }
        }

        // Convert boolean cells
        if (currentCell.type === 'boolean' && typeof currentCell.value === 'number') {
          currentCell.value = currentCell.value === 1;
        }

        const cell: Cell = {
          value: currentCell.value !== undefined ? currentCell.value : '',
          ...(currentCell.type !== undefined && { type: currentCell.type }),
        };

        // If cell has explicit column index, position it correctly; otherwise append
        if (currentCellColIndex !== undefined && currentCellColIndex >= 0) {
          // Ensure cells array is large enough
          while (currentRow.cells.length <= currentCellColIndex) {
            currentRow.cells.push({ value: '', type: 'string' });
          }
          // Track that this column was explicitly set via r attribute
          // This allows us to handle out-of-order cells correctly:
          // - Explicit cells can overwrite empty placeholders (created when padding)
          // - Explicit cells can overwrite other explicit cells (last one wins)
          if (explicitlySetColumns) {
            explicitlySetColumns.add(currentCellColIndex);
          }
          // Place cell at explicit position
          currentRow.cells[currentCellColIndex] = cell;
        } else {
          // No explicit position, append in order
          currentRow.cells.push(cell);
        }
        currentCell = null;
        currentCellColIndex = undefined;
        inInlineStr = false;
        inInlineStrText = false;
        inlineStringBuffer = ''; // Reset buffer after cell is complete
      } else if (event.name === 'v' && inValue) {
        inValue = false;
      } else if (event.name === 't' && inInlineStrText) {
        inInlineStrText = false;
      } else if (event.name === 'r' && inRichTextRun) {
        inRichTextRun = false;
      } else if (event.name === 'rPh' && inRubyPhonetic) {
        inRubyPhonetic = false; // End of pronunciation data
      } else if (event.name === 'is' && inInlineStr) {
        // End of inline string - set accumulated text (even if empty)
        if (currentCell) {
          currentCell.value = inlineStringBuffer;
          currentCell.type = 'string';
        }
        inInlineStr = false;
        inlineStringBuffer = ''; // Reset buffer
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
      } else if (inInlineStrText && currentCell && !inRubyPhonetic) {
        // Inline string text element (<t> inside <is>), but skip if in pronunciation data
        // Accumulate text from multiple rich text runs (multiple <r><t>...</t></r> elements)
        inlineStringBuffer += event.text || '';
      }
    }
  }

  // Yield any remaining row
  if (currentRow && inRow) {
    yield currentRow as Row;
  }
}

