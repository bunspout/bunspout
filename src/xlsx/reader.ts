import type * as yauzl from 'yauzl';
import { openZip, readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import type { ReadOptions } from './types';
import { Workbook, type SheetInfo } from './workbook';

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
  let currentSheet: { name?: string; id?: string; rId?: string } | null = null;
  let inSheet = false;

  for await (const event of parseXmlEvents(readZipEntry(workbookEntry, zipFile))) {
    if (event.type === 'startElement' && event.name === 'sheet') {
      inSheet = true;
      currentSheet = {
        name: event.attributes?.name,
        id: event.attributes?.sheetId,
        rId: event.attributes?.['r:id'],
      };
    } else if (event.type === 'endElement' && event.name === 'sheet' && currentSheet) {
      if (currentSheet.name && currentSheet.id) {
        // Find the corresponding worksheet entry
        const sheetId = parseInt(currentSheet.id, 10);
        const sheetEntry = entries.find(
          (e) => e.fileName === `xl/worksheets/sheet${sheetId}.xml`,
        );
        if (sheetEntry) {
          sheets.push({
            name: currentSheet.name,
            entry: sheetEntry,
          });
        }
      }
      currentSheet = null;
      inSheet = false;
    }
  }

  return sheets;
}

/**
 * Reads an XLSX file and returns a Workbook instance
 */
export async function readXlsx(filePath: string, options?: ReadOptions): Promise<Workbook> {
  // Read file to buffer
  const file = Bun.file(filePath);
  if (!(await file.exists())) {
    throw new Error(`File not found: ${filePath}`);
  }
  const buffer = Buffer.from(await file.arrayBuffer());

  // Open ZIP
  const zipFile = await openZip(buffer);

  // Parse workbook.xml to get sheet names and relationships
  const sheetInfos = await parseWorkbook(zipFile.zipFile, zipFile.entries);

  // Create Workbook instance
  return new Workbook(zipFile, sheetInfos, options);
}

