import { parseSheet } from '@sheet/reader';
import type { ZipFile, ZipEntry } from '@zip/reader';
import { readZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import type { Row } from '../types';
import { parseCoreProperties, parseCustomProperties } from './properties-reader';
import type { SharedStringsCachingStrategy } from './shared-strings-caching';
import { parseSharedStrings } from './shared-strings-reader';
import type { SheetProperties } from './sheet-properties-reader';
import { parseStyles, type StyleFormatMap } from './styles-reader';
import type { ReadOptions, WorkbookProperties } from './types';

export type SheetInfo = {
  name: string;
  entry: ZipEntry;
  properties?: SheetProperties;
};

export class Workbook {
  private zip: ZipFile;
  private readonly _sheets: Sheet[];
  private sharedStrings: SharedStringsCachingStrategy | null = null;
  private styleFormatMap: StyleFormatMap | null = null;
  private readonly options?: ReadOptions;
  private _properties: WorkbookProperties | null = null;

  constructor(zip: ZipFile, sheetInfos: SheetInfo[], options?: ReadOptions) {
    this.zip = zip;
    this.options = options;
    this._sheets = sheetInfos.map(
      (info) => new Sheet(info.name, info.entry, this, info.properties),
    );
  }

  /**
   * Loads workbook properties from core.xml and custom.xml
   */
  async loadProperties(): Promise<void> {
    if (this._properties !== null) {
      return; // Already loaded
    }

    const properties: WorkbookProperties = {};

    // Parse core properties
    // Use case-insensitive matching for robustness across different XLSX file generators
    const coreEntry = this.zip.entries.find(
      (e) => e.fileName.toLowerCase() === 'docprops/core.xml',
    );
    if (coreEntry) {
      const coreProps = await parseCoreProperties(coreEntry, this.zip.zipFile);
      Object.assign(properties, coreProps);
    }

    // Parse custom properties
    // Use case-insensitive matching for robustness across different XLSX file generators
    const customEntry = this.zip.entries.find(
      (e) => e.fileName.toLowerCase() === 'docprops/custom.xml',
    );
    if (customEntry) {
      const customProps = await parseCustomProperties(customEntry, this.zip.zipFile);
      if (Object.keys(customProps).length > 0) {
        properties.customProperties = customProps;
      }
    }

    this._properties = properties;
  }

  /**
   * Gets workbook properties (core and custom properties)
   * Properties are loaded lazily on first access
   */
  async properties(): Promise<WorkbookProperties> {
    await this.loadProperties();
    return this._properties ?? {};
  }

  /**
   * Loads shared strings table if it exists
   */
  async loadSharedStrings(): Promise<void> {
    if (this.sharedStrings !== null) {
      return; // Already loaded
    }

    const sharedStringsEntry = this.zip.entries.find(
      (e) => e.fileName.toLowerCase() === 'xl/sharedstrings.xml',
    );

    if (sharedStringsEntry) {
      this.sharedStrings = await parseSharedStrings(
        sharedStringsEntry,
        this.zip.zipFile,
      );
    } else {
      // Create an empty in-memory strategy when there are no shared strings
      const { InMemoryStrategy } = await import('./shared-strings-caching');
      this.sharedStrings = new InMemoryStrategy(0);
    }
  }

  /**
   * Gets a string from the shared strings table by index
   */
  async getSharedString(index: number): Promise<string | undefined> {
    if (!this.sharedStrings) {
      return undefined;
    }
    return await this.sharedStrings.getString(index);
  }

  /**
   * Loads styles.xml to extract format codes mapped to style indices
   */
  async loadStyles(): Promise<void> {
    if (this.styleFormatMap !== null) {
      return; // Already loaded
    }

    const stylesEntry = this.zip.entries.find(
      (e) => e.fileName.toLowerCase() === 'xl/styles.xml',
    );

    if (stylesEntry) {
      this.styleFormatMap = await parseStyles(stylesEntry, this.zip.zipFile);
    } else {
      // No styles.xml found, create empty map
      this.styleFormatMap = new Map();
    }
  }

  /**
   * Gets the style format map (loaded lazily on first access)
   */
  async getStyleFormatMap(): Promise<StyleFormatMap> {
    await this.loadStyles();
    return this.styleFormatMap ?? new Map();
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
    await this.loadSharedStrings();
    const styleFormatMap = await this.getStyleFormatMap();

    // Stream XML directly from ZIP to parser (no accumulation)
    yield* parseSheet(
      parseXmlEvents(readZipEntry(entry, this.zip.zipFile)),
      (index: number) => this.getSharedString(index),
      this.options,
      styleFormatMap,
    );
  }

  /**
   * Returns all sheets
   */
  sheets(): Sheet[] {
    return [...this._sheets];
  }

  /**
   * Cleans up any temporary resources (e.g., file-based shared strings cache, zip file)
   * Call this when you're done with the workbook to free up disk space and file handles
   */
  async cleanup(): Promise<void> {
    // Clean up shared strings
    if (this.sharedStrings) {
      try {
        await this.sharedStrings.cleanup();
      } catch {
        // Ignore errors during cleanup
      }
      this.sharedStrings = null;
    }

    // Close the zip file
    if (this.zip?.zipFile) {
      try {
        this.zip.zipFile.close();
      } catch {
        // Ignore errors during cleanup
      }
    }
  }

  async [Symbol.asyncDispose]() {
    await this.cleanup();
  }
}

export class Sheet {
  constructor(
    public readonly name: string,
    public readonly entry: ZipEntry,
    private workbook: Workbook,
    private _properties?: SheetProperties,
  ) {}

  /**
   * Returns rows as async iterable (always array-based format)
   */
  async *rows(): AsyncIterable<Row> {
    yield* this.workbook.readSheetRows(this.entry);
  }

  /**
   * Gets sheet properties (column widths, row heights, etc.)
   */
  get properties(): SheetProperties | undefined {
    return this._properties;
  }

  /**
   * Gets default column width for the sheet
   */
  get defaultColumnWidth(): number | undefined {
    return this._properties?.defaultColumnWidth;
  }

  /**
   * Gets column width definitions
   */
  get columnWidths(): SheetProperties['columnWidths'] {
    return this._properties?.columnWidths;
  }

  /**
   * Gets default row height for the sheet
   */
  get defaultRowHeight(): number | undefined {
    return this._properties?.defaultRowHeight;
  }

  /**
   * Gets row height definitions
   */
  get rowHeights(): SheetProperties['rowHeights'] {
    return this._properties?.rowHeights;
  }

  /**
   * Gets whether the sheet is hidden
   */
  get hidden(): boolean {
    return this._properties?.hidden ?? false;
  }
}
