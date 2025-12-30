/**
 * XLSX structure XML generators
 * Keep generators dumb - pass full sheet definitions, not counts
 */

import { escapeXml } from '@utils/xml';

/**
 * Generates [Content_Types].xml
 */
export function generateContentTypes(
  sheets: { id: number }[],
  hasSharedStrings: boolean = false,
  hasCoreProperties: boolean = false,
  hasCustomProperties: boolean = false,
): string {
  const sheetOverrides = sheets
    .map((sheet) => `  <Override PartName="/xl/worksheets/sheet${sheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
    .join('\n');

  const overrides: string[] = [];

  if (hasSharedStrings) {
    overrides.push('  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
  }

  if (hasCoreProperties) {
    overrides.push('  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>');
  }

  if (hasCustomProperties) {
    overrides.push('  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>');
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${overrides.length > 0 ? overrides.join('\n') + '\n' : ''}${sheetOverrides}
</Types>`;
}

/**
 * Generates _rels/.rels
 */
export function generateRels(hasCoreProperties: boolean = false, hasCustomProperties: boolean = false): string {
  const relationships: string[] = [
    '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
  ];

  let idOffset = 1;

  if (hasCoreProperties) {
    relationships.push(
      `  <Relationship Id="rId${idOffset + 1}" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
    );
    idOffset++;
  }

  if (hasCustomProperties) {
    relationships.push(
      `  <Relationship Id="rId${idOffset + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>`,
    );
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.join('\n')}
</Relationships>`;
}

/**
 * Generates xl/workbook.xml
 */
export function generateWorkbook(
  sheets: { name: string; id: number; hidden?: boolean }[],
): string {
  const sheetElements = sheets
    .map(
      (sheet, index) => {
        const stateAttr = sheet.hidden ? ' state="hidden"' : '';
        return `    <sheet name="${escapeXml(sheet.name)}" sheetId="${sheet.id}" r:id="rId${index + 1}"${stateAttr}/>`;
      },
    )
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${sheetElements}
  </sheets>
</workbook>`;
}

/**
 * Generates xl/_rels/workbook.xml.rels
 */
export function generateWorkbookRels(
  sheets: { id: number }[],
  hasSharedStrings: boolean = false,
  hasCoreProperties: boolean = false,
): string {
  const relationships: string[] = [];

  let idOffset = 0;

  // Add shared strings relationship if present
  if (hasSharedStrings) {
    relationships.push(
      '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    );
    idOffset++;
  }

  // Note: Core properties relationship goes in _rels/.rels, not workbook.xml.rels

  // Add sheet relationships (offset IDs based on what's already added)
  sheets.forEach((sheet, index) => {
    relationships.push(
      `  <Relationship Id="rId${index + 1 + idOffset}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheet.id}.xml"/>`,
    );
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.join('\n')}
</Relationships>`;
}

/**
 * Generates docProps/core.xml
 */
export function generateCoreProperties(properties: {
  title?: string | null;
  subject?: string | null;
  application?: string | null;
  creator?: string | null;
  lastModifiedBy?: string | null;
  keywords?: string | null;
  description?: string | null;
  category?: string | null;
  language?: string | null;
  customProperties?: Record<string, string>;
}): string {
  const now = new Date().toISOString();

  const elements: string[] = [];

  if (properties.title) {
    elements.push(`    <dc:title>${escapeXml(properties.title)}</dc:title>`);
  }
  if (properties.subject) {
    elements.push(`    <dc:subject>${escapeXml(properties.subject)}</dc:subject>`);
  }
  if (properties.creator) {
    elements.push(`    <dc:creator>${escapeXml(properties.creator)}</dc:creator>`);
  }
  if (properties.lastModifiedBy) {
    elements.push(`    <cp:lastModifiedBy>${escapeXml(properties.lastModifiedBy)}</cp:lastModifiedBy>`);
  }
  if (properties.keywords) {
    elements.push(`    <cp:keywords>${escapeXml(properties.keywords)}</cp:keywords>`);
  }
  if (properties.description) {
    elements.push(`    <dc:description>${escapeXml(properties.description)}</dc:description>`);
  }
  if (properties.category) {
    elements.push(`    <cp:category>${escapeXml(properties.category)}</cp:category>`);
  }
  if (properties.language) {
    elements.push(`    <dc:language>${escapeXml(properties.language)}</dc:language>`);
  }

  // Always include created and modified dates
  elements.push(`    <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`);
  elements.push(`    <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`);
  elements.push('    <cp:revision>0</cp:revision>');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
${elements.join('\n')}
</cp:coreProperties>`;
}

/**
 * Generates docProps/custom.xml
 */
export function generateCustomProperties(customProperties: Record<string, string>): string {
  if (Object.keys(customProperties).length === 0) {
    return '';
  }

  const properties: string[] = [];
  let pid = 2; // Property IDs start at 2

  for (const [name, value] of Object.entries(customProperties)) {
    properties.push(
      `    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${escapeXml(name)}"><vt:lpwstr>${escapeXml(value)}</vt:lpwstr></property>`,
    );
    pid++;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
${properties.join('\n')}
</Properties>`;
}
