import type * as yauzl from 'yauzl';
import { parseSheet } from '@sheet/reader';
import type { ZipFile, ZipEntry } from '@zip/reader';
import { readZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import type { Row } from '../types';
import { parseSharedStrings } from './shared-strings-reader';

export type SheetInfo = {
  name: string;
  entry: ZipEntry;
};

export class Workbook {
  private zipFile: ZipFile;
  private _sheets: Sheet[];
  private sharedStrings: string[] | null = null;

  constructor(zipFile: ZipFile, sheetInfos: SheetInfo[]) {
    this.zipFile = zipFile;
    this._sheets = sheetInfos.map(
      (info) => new Sheet(info.entry, zipFile.zipFile, info.name, this),
    );
  }

  /**
   * Loads shared strings table if it exists
   */
  async loadSharedStrings(): Promise<void> {
    if (this.sharedStrings !== null) {
      return; // Already loaded
    }

    const sharedStringsEntry = this.zipFile.entries.find(
      (e) => e.fileName === 'xl/sharedStrings.xml',
    );

    if (sharedStringsEntry) {
      this.sharedStrings = await parseSharedStrings(
        sharedStringsEntry,
        this.zipFile.zipFile,
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
   * Returns all sheets
   */
  sheets(): Sheet[] {
    return [...this._sheets];
  }
}

export class Sheet {
  private zipEntry: ZipEntry;
  private zipFile: yauzl.ZipFile;
  public readonly name: string;
  private workbook: Workbook;

  constructor(zipEntry: ZipEntry, zipFile: yauzl.ZipFile, name: string, workbook: Workbook) {
    this.zipEntry = zipEntry;
    this.zipFile = zipFile;
    this.name = name;
    this.workbook = workbook;
  }

  /**
   * Returns rows as async iterable (always array-based format)
   */
  async *rows(): AsyncIterable<Row> {
    // Ensure shared strings are loaded
    await this.workbook.loadSharedStrings();

    // Stream XML directly from ZIP to parser (no accumulation)
    yield* parseSheet(
      parseXmlEvents(readZipEntry(this.zipEntry, this.zipFile)),
      (index: number) => this.workbook.getSharedString(index),
    );
  }
}

