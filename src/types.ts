/*
 * Core domain types for Excel streaming library
 */

// Font style properties for cells
export type FontStyle = {
  /**
   * Bold text
   * @default false
   */
  bold?: boolean;
  /**
   * Italic text
   * @default false
   */
  italic?: boolean;
  /**
   * Underline text
   * @default false
   */
  underline?: boolean;
  /**
   * Strikethrough text
   * @default false
   */
  strikethrough?: boolean;
  /**
   * Font size in points
   * @default 11
   */
  fontSize?: number;
  /**
   * Font color in ARGB format (e.g., "FF0000FF" for blue, "FFFF0000" for red)
   * @default "FF000000" (black)
   */
  fontColor?: string;
  /**
   * Font name
   * @default "Arial"
   */
  fontName?: string;
};

// Cell style definition
export type Style = {
  font?: FontStyle;
};

// Public API Cell type
export type Cell = {
  value: string | number | Date | boolean | null | undefined;
  type?: 'string' | 'number' | 'date' | 'boolean' | 'formula';
  /**
   * Formula string (e.g., "SUM(A1:A5)"). Only present when type is 'formula'.
   * Note: `value` contains the formula with "=" prefix (e.g., "=SUM(A1:A5)").
   */
  formula?: string;
  /**
   * Pre-calculated result from Excel. Only present when type is 'formula'.
   * May be null if the formula result is an error or not available.
   */
  computedValue?: string | number | Date | boolean | null;
  /**
   * Cell style
   */
  style?: Style;
};

// Internal Cell format for XML reader/writer
export type CellResolved = {
  t: 's' | 'n' | 'b' | 'd' | 'e'; // string, number, boolean, date, error
  v: string | number;
};

export type Row = {
  /**
   * Array of cells. May contain undefined/null values for holes (sparse rows).
   * Undefined/null cells are skipped during serialization.
   */
  cells: (Cell | undefined | null)[];
  rowIndex?: number;
  height?: number; // Row height in points
  // Future: styles?: Record<number, Style>;
};

export type XmlEvent = {
  type: 'startElement' | 'endElement' | 'text';
  name?: string;
  attributes?: Record<string, string>;
  text?: string;
};

export type XmlNodeChunk = string | Uint8Array;
