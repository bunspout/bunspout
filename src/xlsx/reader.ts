import type * as yauzl from 'yauzl';
import { openZip, readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import { parseSheetProperties, type SheetProperties } from './sheet-properties-reader';
import type { ReadOptions } from './types';
import { Workbook, type SheetInfo } from './workbook';
import { readFile } from '../adapters';

/**
 * Parses workbook.xml to extract sheet information
 */
async function parseWorkbook(
  zipFile: yauzl.ZipFile,
  entries: ZipEntry[],
): Promise<SheetInfo[]> {
  // Find workbook.xml entry
  const workbookEntry = entries.find((e) => e.fileName === 'xl/workbook.xml');
  if (!workbookEntry) {
    throw new Error('workbook.xml not found in XLSX file');
  }

  // Stream XML directly from ZIP to parser (no accumulation)
  const sheets: SheetInfo[] = [];
  let currentSheet: { name?: string; id?: string; rId?: string; hidden?: boolean } | null = null;

  for await (const event of parseXmlEvents(readZipEntry(workbookEntry, zipFile))) {
    if (event.type === 'startElement' && event.name === 'sheet') {
      const state = event.attributes?.state;
      currentSheet = {
        name: event.attributes?.name,
        id: event.attributes?.sheetId,
        rId: event.attributes?.['r:id'],
        hidden: state === 'hidden',
      };
    } else if (event.type === 'endElement' && event.name === 'sheet' && currentSheet) {
      if (currentSheet.name && currentSheet.id) {
        // Find the corresponding worksheet entry
        const sheetId = parseInt(currentSheet.id, 10);
        const sheetEntry = entries.find(
          (e) => e.fileName === `xl/worksheets/sheet${sheetId}.xml`,
        );
        if (sheetEntry) {
          // Parse sheet properties (column widths, row heights, etc.)
          let sheetProperties: SheetProperties | undefined;
          try {
            sheetProperties = await parseSheetProperties(sheetEntry, zipFile);
          } catch {
            // If parsing fails, continue without properties
            sheetProperties = undefined;
          }

          // Add hidden state to properties (hidden comes from workbook.xml, not worksheet XML)
          if (currentSheet.hidden !== undefined) {
            if (!sheetProperties) {
              sheetProperties = {};
            }
            sheetProperties.hidden = currentSheet.hidden;
          }

          sheets.push({
            name: currentSheet.name,
            entry: sheetEntry,
            properties: sheetProperties,
          });
        }
      }
      currentSheet = null;
    }
  }

  return sheets;
}

/**
 * Reads an XLSX file and returns a Workbook instance
 */
export async function readXlsx(filePath: string, options?: ReadOptions): Promise<Workbook> {
  // Read file to buffer
  const buffer = await readFile(filePath);

  // Open ZIP
  const zipFile = await openZip(buffer);

  // Parse workbook.xml to get sheet names and relationships
  const sheetInfos = await parseWorkbook(zipFile.zipFile, zipFile.entries);

  // Create Workbook instance
  const workbook = new Workbook(zipFile, sheetInfos, options);

  // Load properties asynchronously (lazy loading)
  // Properties will be loaded when accessed via workbook.properties
  // This avoids blocking on properties parsing during workbook creation

  return workbook;
}
