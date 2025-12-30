/*
 * Shared strings table builder and generator
 */

import { escapeXml } from '@utils/xml';

/**
 * Shared strings table builder
 */
export class SharedStringsTable {
  private strings: string[] = [];
  private indexMap: Map<string, number> = new Map();
  private totalCount: number = 0;

  /**
   * Adds a string to the table and returns its index
   */
  addString(str: string): number {
    this.totalCount++;
    if (this.indexMap.has(str)) {
      return this.indexMap.get(str)!;
    }
    const index = this.strings.length;
    this.strings.push(str);
    this.indexMap.set(str, index);
    return index;
  }

  /**
   * Gets the index of a string (must have been added first)
   */
  getIndex(str: string): number {
    const index = this.indexMap.get(str);
    if (index === undefined) {
      throw new Error(`String "${str}" not found in shared strings table`);
    }
    return index;
  }

  /**
   * Gets all strings in order
   */
  getStrings(): readonly string[] {
    return this.strings;
  }

  /**
   * Gets the count of unique strings
   */
  getUniqueCount(): number {
    return this.strings.length;
  }

  /**
   * Gets the total count of string references
   */
  getTotalCount(): number {
    return this.totalCount;
  }

  /**
   * Generates the shared strings XML
   */
  generateXml(): string {
    if (this.strings.length === 0) {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`;
    }

    const stringElements = this.strings
      .map((str) => {
        const escaped = escapeXml(str);
        return `    <si><t>${escaped}</t></si>`;
      })
      .join('\n');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.totalCount}" uniqueCount="${this.strings.length}">
${stringElements}
</sst>`;
  }
}
