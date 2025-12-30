import type { Cell, CellResolved } from '../types';

/**
 * Converts a Cell to CellResolved format for XML serialization
 */
export function resolveCell(cell: Cell): CellResolved {
  // Handle empty/null cells
  if (cell.value === null || cell.value === undefined || cell.value === '') {
    return { t: 's', v: '' };
  }

  // Use explicit type if provided
  if (cell.type === 'date') {
    return { t: 'd', v: cell.value as number };
  }
  if (cell.type === 'boolean') {
    return { t: 'b', v: cell.value as number };
  }
  if (cell.type === 'string') {
    return { t: 's', v: cell.value as string };
  }
  if (cell.type === 'number') {
    return { t: 'n', v: cell.value as number };
  }

  // Auto-detect from value type
  if (typeof cell.value === 'string') {
    return { t: 's', v: cell.value };
  }
  if (typeof cell.value === 'number') {
    return { t: 'n', v: cell.value };
  }

  // Fallback
  return { t: 's', v: String(cell.value) };
}
