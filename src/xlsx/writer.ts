import { collect } from '@utils/transforms';
import { createZipWriter, writeZipEntry, endZipWriter } from '@zip/writer';
import { resolveCell } from '@xml/cell-resolver';
import { writeSheetXml } from '@xml/writer';
import { SharedStringsTable } from './shared-strings';
import { generateContentTypes, generateRels, generateWorkbook, generateWorkbookRels, generateCoreProperties, generateCustomProperties } from './structure';
import type { WorkbookDefinition, WriterOptions } from './types';
import { stringToBytes } from '../adapters/common';
import type { Row } from '../types';

/**
 * Builds shared strings table from all rows
 */
function buildSharedStringsTable(allRows: Row[]): SharedStringsTable {
  const table = new SharedStringsTable();
  for (const row of allRows) {
    for (const cell of row.cells) {
      if (cell !== undefined && cell !== null) {
        const resolved = resolveCell(cell);
        if (resolved.t === 's') {
          table.addString(resolved.v as string);
        }
      }
    }
  }
  return table;
}

/**
 * Writes an XLSX file from a workbook definition
 */
export async function writeXlsx(
  filePath: string,
  definition: WorkbookDefinition,
  options?: WriterOptions,
): Promise<void> {
  const opts = {
    sharedStrings: 'inline' as const,
    ...options,
  };

  const zipWriter = createZipWriter();

  // Write each sheet
  const sheetInfos = definition.sheets.map((sheetDef, index) => ({
    name: sheetDef.name,
    id: index + 1,
    hidden: sheetDef.hidden ?? false,
  }));

  let sharedStringsTable: SharedStringsTable | null = null;
  const allSheetRows: Row[][] = [];

  // If using shared strings, collect all rows first to build the table
  if (opts.sharedStrings === 'shared') {
    for (let i = 0; i < definition.sheets.length; i++) {
      const sheetDef = definition.sheets[i]!;
      const rows = await collect<Row>(sheetDef.rows);
      allSheetRows.push(rows);
    }
    sharedStringsTable = buildSharedStringsTable(allSheetRows.flat());
  }

  // Write each sheet
  for (let i = 0; i < definition.sheets.length; i++) {
    const sheetDef = definition.sheets[i]!;
    const sheetId = i + 1;

    // Prepare column width options
    const columnWidthOptions = {
      defaultColumnWidth: sheetDef.defaultColumnWidth,
      columnWidths: sheetDef.columnWidths,
      autoDetectColumnWidth: sheetDef.autoDetectColumnWidth,
    };

    // Prepare row height options
    const rowHeightOptions = {
      defaultRowHeight: sheetDef.defaultRowHeight,
      rowHeights: sheetDef.rowHeights,
    };

    // Generate sheet XML
    let sheetXml: AsyncIterable<string>;
    if (opts.sharedStrings === 'shared' && sharedStringsTable) {
      // Use collected rows with shared strings
      const rows = allSheetRows[i]!;
      const getStringIndex = (str: string) => sharedStringsTable!.getIndex(str);
      sheetXml = writeSheetXml(
        (async function* () {
          for (const row of rows) yield row;
        })(),
        {
          getStringIndex,
          columnWidths: columnWidthOptions,
          ...rowHeightOptions,
        },
      );
    } else {
      // Use original rows with inline strings
      sheetXml = writeSheetXml(sheetDef.rows, {
        columnWidths: columnWidthOptions,
        ...rowHeightOptions,
      });
    }
    const xmlStrings = await collect<string>(sheetXml);

    // Write sheet XML to ZIP
    await writeZipEntry(
      zipWriter,
      `xl/worksheets/sheet${sheetId}.xml`,
      stringToBytes(xmlStrings.join('')),
    );
  }

  // Write shared strings if enabled
  if (opts.sharedStrings === 'shared' && sharedStringsTable) {
    await writeZipEntry(
      zipWriter,
      'xl/sharedStrings.xml',
      stringToBytes(sharedStringsTable.generateXml()),
    );
  }

  // Write core properties if provided
  const hasProperties = definition.properties !== undefined;
  const hasCustomProps = definition.properties?.customProperties !== undefined &&
                         Object.keys(definition.properties.customProperties || {}).length > 0;

  if (hasProperties && definition.properties) {
    await writeZipEntry(
      zipWriter,
      'docProps/core.xml',
      stringToBytes(generateCoreProperties(definition.properties)),
    );
  }

  // Write custom properties if provided
  if (hasCustomProps && definition.properties?.customProperties) {
    const customXml = generateCustomProperties(definition.properties.customProperties);
    if (customXml) {
      await writeZipEntry(
        zipWriter,
        'docProps/custom.xml',
        stringToBytes(customXml),
      );
    }
  }

  // Write structure files
  await writeZipEntry(
    zipWriter,
    '[Content_Types].xml',
    stringToBytes(generateContentTypes(sheetInfos, opts.sharedStrings === 'shared', hasProperties, hasCustomProps)),
  );

  await writeZipEntry(
    zipWriter,
    '_rels/.rels',
    stringToBytes(generateRels(hasProperties, hasCustomProps)),
  );

  await writeZipEntry(
    zipWriter,
    'xl/workbook.xml',
    stringToBytes(generateWorkbook(sheetInfos)),
  );

  await writeZipEntry(
    zipWriter,
    'xl/_rels/workbook.xml.rels',
    stringToBytes(generateWorkbookRels(sheetInfos, opts.sharedStrings === 'shared', hasProperties)),
  );

  // Write ZIP to file
  const buffer = await endZipWriter(zipWriter);
  await Bun.write(filePath, buffer);
}
