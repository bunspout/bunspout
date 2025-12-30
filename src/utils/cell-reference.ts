/*
 * Excel cell reference utilities
 * Handles conversion between Excel column letters (A, B, ..., Z, AA, AB, ...)
 * and 0-based column indices, as well as parsing/generating cell references.
 */

/**
 * Converts a column index (0-based) to Excel column letter (A, B, ..., Z, AA, AB, ...)
 * @param colIndex - Column index (0-based: 0 = A, 1 = B, etc.)
 * @returns Excel column letter(s)
 */
export function columnIndexToLetter(colIndex: number): string {
  let result = '';
  colIndex++; // Convert to 1-based
  while (colIndex > 0) {
    colIndex--;
    result = String.fromCharCode(65 + (colIndex % 26)) + result;
    colIndex = Math.floor(colIndex / 26);
  }
  return result;
}

/**
 * Parses Excel column letter(s) to 0-based column index
 * @param colLetter - Excel column letter (A, B, ..., Z, AA, AB, ...)
 * @returns 0-based column index (0 = A, 1 = B, etc.)
 */
export function columnLetterToIndex(colLetter: string): number {
  let index = 0;
  for (let i = 0; i < colLetter.length; i++) {
    index = index * 26 + (colLetter.charCodeAt(i) - 64); // 'A' = 65, so -64 gives 1
  }
  return index - 1; // Convert to 0-based
}

/**
 * Generates a cell reference (e.g., "A1", "B2") from row and column indices
 * @param rowIndex - Row index (1-based: 1 = first row)
 * @param colIndex - Column index (0-based: 0 = A, 1 = B, etc.)
 * @returns Excel cell reference string
 */
export function getCellReference(rowIndex: number, colIndex: number): string {
  return `${columnIndexToLetter(colIndex)}${rowIndex}`;
}

/**
 * Parses cell reference (e.g., "A1", "B2") to row and column indices
 * @param cellRef - Cell reference string
 * @returns Object with rowIndex (1-based) and colIndex (0-based), or null if invalid
 */
export function parseCellReference(cellRef: string): { rowIndex: number; colIndex: number } | null {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const colLetter = match[1]!;
  const rowNum = parseInt(match[2]!, 10);
  return {
    rowIndex: rowNum,
    colIndex: columnLetterToIndex(colLetter),
  };
}
