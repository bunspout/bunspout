import { parseSheet } from '@sheet/reader';
import type { ZipFile, ZipEntry } from '@zip/reader';
import { readZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import type { Row } from '../types';
import { parseSharedStrings } from './shared-strings-reader';
import type { ReadOptions } from './types';

export type SheetInfo = {
  name: string;
  entry: ZipEntry;
};

export class Workbook {
  private zip: ZipFile;
  private _sheets: Sheet[];
  private sharedStrings: string[] | null = null;
  private options?: ReadOptions;

  constructor(zip: ZipFile, sheetInfos: SheetInfo[], options?: ReadOptions) {
    this.zip = zip;
    this.options = options;
    this._sheets = sheetInfos.map(
      (info) => new Sheet(info.name, info.entry, this),
    );
  }

  /**
   * Loads shared strings table if it exists
   */
  async loadSharedStrings(): Promise<void> {
    if (this.sharedStrings !== null) {
      return; // Already loaded
    }

    const sharedStringsEntry = this.zip.entries.find(
      (e) => e.fileName === 'xl/sharedStrings.xml',
    );

    if (sharedStringsEntry) {
      this.sharedStrings = await parseSharedStrings(
        sharedStringsEntry,
        this.zip.zipFile,
      );
    } else {
      this.sharedStrings = [];
    }
  }

  /**
   * Gets a string from the shared strings table by index
   */
  getSharedString(index: number): string | undefined {
    if (!this.sharedStrings) {
      return undefined;
    }
    return this.sharedStrings[index];
  }

  /**
   * Gets a sheet by name or index (synchronous - metadata already loaded)
   */
  sheet(nameOrIndex: string | number): Sheet {
    if (typeof nameOrIndex === 'string') {
      const sheet = this._sheets.find((s) => s.name === nameOrIndex);
      if (!sheet) {
        throw new Error(`Sheet "${nameOrIndex}" not found`);
      }
      return sheet;
    } else {
      const sheet = this._sheets[nameOrIndex];
      if (!sheet) {
        throw new Error(`Sheet at index ${nameOrIndex} not found`);
      }
      return sheet;
    }
  }

  /**
   * Reads rows from a sheet entry
   */
  async *readSheetRows(entry: ZipEntry): AsyncIterable<Row> {
    // Ensure shared strings are loaded
    await this.loadSharedStrings();

    // Stream XML directly from ZIP to parser (no accumulation)
    yield* parseSheet(
      parseXmlEvents(readZipEntry(entry, this.zip.zipFile)),
      (index: number) => this.getSharedString(index),
      this.options,
    );
  }

  /**
   * Returns all sheets
   */
  sheets(): Sheet[] {
    return [...this._sheets];
  }
}

export class Sheet {
  constructor(
    public readonly name: string,
    public readonly entry: ZipEntry,
    private workbook: Workbook,
  ) {}

  /**
   * Returns rows as async iterable (always array-based format)
   */
  async *rows(): AsyncIterable<Row> {
    yield* this.workbook.readSheetRows(this.entry);
  }
}

