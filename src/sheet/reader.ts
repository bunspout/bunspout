import { format } from 'date-fns';
import { getFormatCodeForStyle, type StyleFormatMap } from '@xlsx/styles-reader';
import type { ReadOptions } from '@xlsx/types';
import { parseCellReference } from '@utils/cell-reference';
import { convertExcelTimestamp } from '@utils/dates';
import {
  convertExcelFormatToDateFns,
  DEFAULT_DATE_FORMAT,
  isDateFormatCode,
} from '@utils/format-codes';
import type { Cell, Row, XmlEvent } from '../types';

/**
 * Checks if a row is empty (no cells or all cells are empty strings/undefined)
 * Note: null values are considered valid (invalid but present data), not empty
 */
function isEmptyRow(row: Partial<Row> | null): boolean {
  if (!row || !row.cells) return true;
  if (row.cells.length === 0) return true;
  return row.cells.every(cell => !cell || cell.value === '' || cell.value === undefined);
}

/**
 * Determines if a row should be yielded based on skipEmptyRows option
 */
function shouldYieldRow(row: Partial<Row> | null, options?: ReadOptions): boolean {
  if (!row) return false;
  const isEmpty = isEmptyRow(row);
  const shouldSkip = options?.skipEmptyRows !== false; // Default to true
  return !isEmpty || !shouldSkip;
}

/**
 * Finalizes a cell by applying semantic post-processing:
 * - Inline string finalization
 * - Defaulting empty cells
 * - Date detection and formatting
 * - Boolean coercion
 *
 * This is separate from parsing logic for testability and clarity.
 */
function finalizeCell(
  cell: Partial<Cell>,
  context: {
    options?: ReadOptions;
    styleFormatMap?: StyleFormatMap;
    currentCellStyleIndex?: number;
    inlineStringBuffer?: string;
  },
): Cell {
  const {
    options,
    styleFormatMap,
    currentCellStyleIndex,
    inlineStringBuffer = '',
  } = context;

  // Handle inline string finalization
  if (cell.value === undefined && inlineStringBuffer !== undefined) {
    cell.value = inlineStringBuffer;
    cell.type = 'string';
  }

  // Default empty cells
  if (cell.value === undefined) {
    cell.value = '';
    cell.type = 'string';
  }

  // Handle date cells
  if (options) {
    // Check if this is an ISO 8601 date string (t="d")
    if (cell.type === 'date' && typeof cell.value === 'string') {
      // ISO 8601 date string - parse to Date object if shouldFormatDates is false
      if (!options.shouldFormatDates) {
        try {
          cell.value = new Date(cell.value);
        } catch {
          // If parsing fails, keep as string
        }
      }
      // If shouldFormatDates is true, keep as string (already formatted)
    } else if (typeof cell.value === 'number') {
      // Check if this numeric cell should be treated as a date
      let isDate = cell.type === 'date';
      let formatCode: string | null = null;

      // Check format code to detect dates
      if (styleFormatMap && currentCellStyleIndex !== undefined) {
        formatCode = getFormatCodeForStyle(currentCellStyleIndex, styleFormatMap);
        if (formatCode && isDateFormatCode(formatCode)) {
          isDate = true;
          // Override the type if it was automatically set to 'number'
          cell.type = 'date';
        }
      }

      // Convert numeric date cells
      if (isDate) {
        try {
          const date = convertExcelTimestamp(cell.value, options.use1904Dates ?? false);

          if (options.shouldFormatDates && formatCode) {
            // Format the date according to the format code
            const dateFnsFormat = convertExcelFormatToDateFns(formatCode);
            cell.value = format(date, dateFnsFormat);
          } else if (options.shouldFormatDates) {
            // No format code found, use default format
            cell.value = format(date, DEFAULT_DATE_FORMAT);
          } else {
            // Return Date object
            cell.value = date;
          }
        } catch {
          // If conversion fails, return null for invalid dates
          cell.value = null;
        }
      }
    }
  }

  // Convert boolean cells
  if (cell.type === 'boolean' && typeof cell.value === 'number') {
    cell.value = cell.value === 1;
  }

  // Build final cell object
  return {
    value: cell.value !== undefined ? cell.value : '',
    ...(cell.type !== undefined && { type: cell.type }),
    ...(cell.formula !== undefined && { formula: cell.formula }),
    ...(cell.computedValue !== undefined && { computedValue: cell.computedValue }),
  };
}

/**
 * Parses a sheet from XML events, yielding rows
 *
 * Note: This parser implements a streaming state machine over SpreadsheetML.
 * State is tracked via flags (inRow, inCell, etc.) and buffers.
 * Do not refactor into callbacks or recursion.
 */
export async function* parseSheet(
  xmlEvents: AsyncIterable<XmlEvent>,
  getSharedString?: (index: number) => Promise<string | undefined>,
  options?: ReadOptions,
  styleFormatMap?: StyleFormatMap,
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
  let inFormula = false; // Track <f> elements (formula)
  let expectedColumnCount: number | null = null; // From spans attribute (1-based last column)
  let currentCellColIndex: number | undefined = undefined; // Column index from cell r attribute
  let explicitlySetColumns: Set<number> | null = null; // Track which column indices have been explicitly set via r attribute
  let inlineStringBuffer: string = ''; // Accumulate text from multiple rich text runs in inline strings
  let formulaBuffer: string = ''; // Accumulate formula text from <f> element
  let computedValueBuffer: string = ''; // Accumulate computed value from <v> element in formula cells
  let originalCellType: string | undefined = undefined; // Track original t attribute for formula cells
  let currentCellStyleIndex: number | undefined = undefined; // Style index from cell s attribute

  for await (const event of xmlEvents) {
    // ============================================================
    // START ELEMENT: Initialize state for new XML elements
    // ============================================================
    if (event.type === 'startElement') {
      // ROW START
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
      }
      // CELL START
      else if (event.name === 'c' && inRow) {
        inCell = true;
        const cellType = event.attributes?.t;
        const cellRef = event.attributes?.r; // Cell reference like "A1", "B1", etc.
        const styleIndexAttr = event.attributes?.s; // Style index
        currentCellColIndex = cellRef ? parseCellReference(cellRef)?.colIndex : undefined;
        originalCellType = cellType; // Store original type attribute for formula cells

        // Parse style index
        if (styleIndexAttr) {
          const parsed = parseInt(styleIndexAttr, 10);
          currentCellStyleIndex = !isNaN(parsed) ? parsed : undefined;
        } else {
          currentCellStyleIndex = undefined;
        }
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
        inFormula = false;
        inlineStringBuffer = ''; // Reset text accumulation buffer
        formulaBuffer = ''; // Reset formula buffer
        computedValueBuffer = ''; // Reset computed value buffer
      }
      // FORMULA START
      else if (event.name === 'f' && inCell) {
        inFormula = true;
        formulaBuffer = ''; // Initialize formula buffer
      }
      // VALUE START
      else if (event.name === 'v' && inCell && !inInlineStr) {
        inValue = true;
        // If this is a formula cell, accumulate the computed value separately
        // Note: <f> always comes before <v> in XLSX, so inFormula is already false here
        if (currentCell?.type === 'formula') {
          computedValueBuffer = ''; // Initialize computed value buffer
        }
      }
      // INLINE STRING START
      else if (event.name === 'is' && inCell) {
        inInlineStr = true;
        inlineStringBuffer = ''; // Initialize text accumulation buffer for inline string
      }
      // RICH TEXT RUN START
      else if (event.name === 'r' && inInlineStr) {
        inRichTextRun = true;
      }
      // TEXT ELEMENT START (for inline strings)
      else if (event.name === 't' && !inRubyPhonetic) {
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
      }
      // RUBY PHONETIC START (skip pronunciation data)
      else if (event.name === 'rPh' && (inInlineStr || inRichTextRun)) {
        inRubyPhonetic = true; // Skip pronunciation data
      }
    }

    // ============================================================
    // TEXT: Accumulate text content from XML text nodes
    // ============================================================
    else if (event.type === 'text') {
      // FORMULA TEXT
      if (inFormula && currentCell) {
        formulaBuffer += event.text || '';
      }
      // VALUE TEXT
      else if (inValue && currentCell && !inInlineStr) {
        const text = event.text || '';

        // If this is a formula cell, accumulate the computed value
        if (currentCell.type === 'formula') {
          computedValueBuffer += text;
        } else if (currentCell.type === 'string' && getSharedString) {
          // If cell type is 's' (shared string), look up the string by index
          const index = parseInt(text, 10);
          if (!isNaN(index)) {
            const sharedString = await getSharedString(index);
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
      }
      // INLINE STRING TEXT
      else if (inInlineStrText && currentCell && !inRubyPhonetic) {
        // Accumulate text from multiple rich text runs (multiple <r><t>...</t></r> elements)
        inlineStringBuffer += event.text || '';
      }
    }

    // ============================================================
    // END ELEMENT: Finalize and yield completed structures
    // ============================================================
    else if (event.type === 'endElement') {
      // ROW END (yield here)
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

        if (shouldYieldRow(currentRow, options)) {
          yield currentRow as Row;
        }

        currentRow = null;
        expectedColumnCount = null;
        explicitlySetColumns = null;
      }
      // CELL END (finalizeCell here)
      else if (event.name === 'c' && inCell && currentCell && currentRow && currentRow.cells) {
        inCell = false;

        // Finalize cell with semantic post-processing
        const cell = finalizeCell(currentCell, {
          options,
          styleFormatMap,
          currentCellStyleIndex,
          inlineStringBuffer: inInlineStr ? inlineStringBuffer : undefined,
        });

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
        currentCellStyleIndex = undefined; // Reset style index
        originalCellType = undefined; // Reset original cell type
        inInlineStr = false;
        inInlineStrText = false;
        inFormula = false;
        inlineStringBuffer = ''; // Reset buffer after cell is complete
        formulaBuffer = ''; // Reset formula buffer
        computedValueBuffer = ''; // Reset computed value buffer
      }
      // FORMULA END
      else if (event.name === 'f' && inFormula) {
        inFormula = false;
        if (currentCell) {
          currentCell.type = 'formula';
          // Store formula without "=" prefix in formula property
          currentCell.formula = formulaBuffer;
          // Store formula with "=" prefix in value property
          currentCell.value = `=${formulaBuffer}`;
        }
        formulaBuffer = ''; // Reset buffer
      }
      // VALUE END
      else if (event.name === 'v' && inValue) {
        inValue = false;
        // If this was a formula cell, we've finished reading the computed value
        if (currentCell && currentCell?.type === 'formula') {
          const text = computedValueBuffer.trim();
          // Note: Formula computed values are typically stored directly in <v>,
          // not as shared string references. If a formula cell has t="s", it would
          // indicate the computed value type, but Excel usually stores the actual value.
          // For now, we parse the computed value directly. If shared string support
          // is needed for formula computed values, we can add it later.

          // Check if the original cell type was boolean (t="b")
          // Formula cells with boolean results are stored as 1 (TRUE) or 0 (FALSE)
          const isNumeric = text !== '' && !isNaN(Number(text));
          if (originalCellType === 'b' && isNumeric) {
            currentCell.computedValue = Number(text) === 1;
          } else if (isNumeric) {
            // Try to parse as number if it looks like a number
            currentCell.computedValue = Number(text);
          } else {
            // Store as string or null
            currentCell.computedValue = text === '' ? null : text;
          }
          computedValueBuffer = ''; // Reset buffer
        }
      }
      // TEXT ELEMENT END
      else if (event.name === 't' && inInlineStrText) {
        inInlineStrText = false;
      }
      // RICH TEXT RUN END
      else if (event.name === 'r' && inRichTextRun) {
        inRichTextRun = false;
      }
      // RUBY PHONETIC END
      else if (event.name === 'rPh' && inRubyPhonetic) {
        inRubyPhonetic = false; // End of pronunciation data
      }
      // INLINE STRING END
      else if (event.name === 'is' && inInlineStr) {
        // End of inline string - set accumulated text (even if empty)
        if (currentCell) {
          currentCell.value = inlineStringBuffer;
          currentCell.type = 'string';
        }
        inInlineStr = false;
        inlineStringBuffer = ''; // Reset buffer
      }
    }
  }

  // Yield any remaining row
  if (currentRow && inRow && shouldYieldRow(currentRow, options)) {
    yield currentRow as Row;
  }
}
