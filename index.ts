/**
 * Bunspout - Fast, streaming Excel read/write library
 */

// High-level XLSX API
export { writeXlsx } from './src/xlsx/writer';
export { readXlsx } from './src/xlsx/reader';
export { Workbook, Sheet } from './src/xlsx/workbook';
export type { WorkbookDefinition, SheetDefinition, WriterOptions, WorkbookProperties } from './src/xlsx/types';

// Cell and Row factories
export { cell, cellFromString, cellFromNumber, cellFromDate, cellFromBoolean, cellFromNull } from './src/sheet/cell';
export { row } from './src/sheet/row';
export type { RowOptions } from './src/sheet/row';

// Core types
export type { Cell, Row, CellResolved } from './src/types';

// Utility functions
export { mapRows, filterRows, limitRows, collect } from './src/utils/transforms';
