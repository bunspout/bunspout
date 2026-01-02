import { collect } from '@utils/transforms';
import { createZipWriter, writeZipEntry, endZipWriter } from '@zip/writer';
import { resolveCell } from '@xml/cell-resolver';
import { writeSheetXml } from '@xml/writer';
import { SharedStringsTable } from './shared-strings';
import { generateContentTypes, generateRels, generateWorkbook, generateWorkbookRels, generateCoreProperties, generateCustomProperties } from './structure';
import { StyleRegistry } from './styles';
import type { WorkbookDefinition, WriterOptions } from './types';
import { sheetNameSchema, workbookPropertiesSchema } from './validation';
import { writeFile } from '../adapters';
import { stringToBytes } from '../adapters';
import type { Row, Style } from '../types';

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
 * @throws {z.ZodError} If validation fails (sheet names, properties, etc.)
 */
export async function writeXlsx(
  filePath: string,
  definition: WorkbookDefinition,
  options?: WriterOptions,
): Promise<void> {
  // Validate sheet names
  for (const sheet of definition.sheets) {
    sheetNameSchema.parse(sheet.name);
  }

  // Validate workbook properties if provided
  if (definition.properties) {
    workbookPropertiesSchema.parse(definition.properties);
  }

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
  let styleRegistry: StyleRegistry | null = null;
  const allSheetRows: Row[][] = [];

  // Collect rows only if shared strings are enabled (styles are registered incrementally)
  // Note: We build shared strings before styles are registered. This is safe because:
  // - String values are currently independent of styles (no number formats, locale formatting)
  // - If we later add format-dependent stringification (number formats, locale), shared strings
  //   may need to be built after style registration or become format-aware.
  if (opts.sharedStrings === 'shared') {
    // Collect all rows for shared strings table
    for (let i = 0; i < definition.sheets.length; i++) {
      const sheetDef = definition.sheets[i]!;
      const rows = await collect<Row>(sheetDef.rows);
      allSheetRows.push(rows);
    }
    sharedStringsTable = buildSharedStringsTable(allSheetRows.flat());
  }

  // Create style registry for incremental style registration during sheet writing
  styleRegistry = new StyleRegistry();

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
    // Styles are registered incrementally as cells are written (preserves streaming)
    const getStringIndex = opts.sharedStrings === 'shared' && sharedStringsTable
      ? (str: string) => sharedStringsTable!.getIndex(str)
      : undefined;
    // getStyleIndex is always provided (StyleRegistry is always created)
    // It's only called when a cell has a style, and returns the cellXfs index (>= 1)
    const getStyleIndex = (style: Style) => {
      // Register style incrementally and return its cellXfs index
      // StyleRegistry handles the offset internally (index 0 is reserved for default style)
      return styleRegistry!.addStyle(style);
    };

    // Use collected rows if shared strings enabled, otherwise stream original rows
    const rowsToWrite = opts.sharedStrings === 'shared' && allSheetRows[i]
      ? (async function* () {
        for (const row of allSheetRows[i]!) yield row;
      })()
      : sheetDef.rows;

    const sheetXml = writeSheetXml(rowsToWrite, {
      getStringIndex,
      getStyleIndex,
      columnWidths: columnWidthOptions,
      ...rowHeightOptions,
    });
    // Note: We stream row computation (process rows one at a time), but buffer the XML output
    // before writing to ZIP. This is acceptable for now, but means full sheet XML is in memory.
    // Future optimization: stream XML chunks directly to ZIP without full buffering.
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

  // Write styles.xml if any styles were registered (generated from registry, not all rows)
  if (styleRegistry.getCount() > 0) {
    await writeZipEntry(
      zipWriter,
      'xl/styles.xml',
      stringToBytes(styleRegistry.generateXml()),
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
  const hasStyles = styleRegistry.getCount() > 0;
  const hasSharedStrings = opts.sharedStrings === 'shared';

  // Calculate relationship ID offset for sheets (sheets come after shared strings and styles)
  let idOffset = 0;
  if (hasSharedStrings) idOffset++;
  if (hasStyles) idOffset++;
  await writeZipEntry(
    zipWriter,
    '[Content_Types].xml',
    stringToBytes(generateContentTypes(sheetInfos, hasSharedStrings, hasProperties, hasCustomProps, hasStyles)),
  );

  await writeZipEntry(
    zipWriter,
    '_rels/.rels',
    stringToBytes(generateRels(hasProperties, hasCustomProps)),
  );

  await writeZipEntry(
    zipWriter,
    'xl/workbook.xml',
    stringToBytes(generateWorkbook(sheetInfos, idOffset)),
  );

  await writeZipEntry(
    zipWriter,
    'xl/_rels/workbook.xml.rels',
    stringToBytes(generateWorkbookRels(sheetInfos, hasSharedStrings, hasStyles)),
  );

  // Write ZIP to file
  const buffer = await endZipWriter(zipWriter);
  await writeFile(filePath, buffer);
}
