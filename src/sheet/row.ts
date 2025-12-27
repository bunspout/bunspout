import type { Cell, Row } from './types';

export interface RowOptions {
  rowIndex?: number;
  // Future: styles?: Record<number, Style>;
}

/**
 * Creates a row from cells with optional options
 */
export function row(cells: Cell[], options?: RowOptions): Row {
  return {
    cells,
    rowIndex: options?.rowIndex,
  };
}

