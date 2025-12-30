import type * as yauzl from 'yauzl';
import { getBuiltInFormatCode } from '@utils/format-codes';
import { readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';

/**
 * Map from style index to format code
 */
export type StyleFormatMap = Map<number, string>;

/**
 * Parses styles.xml to extract format codes mapped to style indices
 */
export async function parseStyles(
  zipEntry: ZipEntry,
  zipFile: yauzl.ZipFile,
): Promise<StyleFormatMap> {
  const formatMap = new Map<number, string>(); // numFmtId -> formatCode
  const styleMap = new Map<number, number>(); // styleIndex -> numFmtId
  let inNumFmts = false;
  let inCellXfs = false;
  let currentNumFmtId: number | null = null;
  let currentFormatCode: string = '';
  let currentStyleIndex = 0;
  let currentNumFmtIdAttr: string | undefined = undefined;

  for await (const event of parseXmlEvents(readZipEntry(zipEntry, zipFile))) {
    if (event.type === 'startElement') {
      if (event.name === 'numFmts') {
        inNumFmts = true;
      } else if (event.name === 'numFmt' && inNumFmts) {
        const numFmtIdAttr = event.attributes?.['numFmtId'];
        const formatCodeAttr = event.attributes?.['formatCode'];
        if (numFmtIdAttr) {
          const parsed = parseInt(numFmtIdAttr, 10);
          if (!isNaN(parsed)) {
            currentNumFmtId = parsed;
            currentFormatCode = formatCodeAttr || '';
          }
        }
      } else if (event.name === 'cellXfs') {
        inCellXfs = true;
        currentStyleIndex = 0;
      } else if (event.name === 'xf' && inCellXfs) {
        // Get numFmtId from xf element
        currentNumFmtIdAttr = event.attributes?.['numFmtId'];
      }
    } else if (event.type === 'endElement') {
      if (event.name === 'numFmts') {
        inNumFmts = false;
      } else if (event.name === 'numFmt' && currentNumFmtId !== null) {
        // Store format code for this numFmtId
        if (currentFormatCode) {
          formatMap.set(currentNumFmtId, currentFormatCode);
        }
        currentNumFmtId = null;
        currentFormatCode = '';
      } else if (event.name === 'cellXfs') {
        inCellXfs = false;
      } else if (event.name === 'xf' && inCellXfs) {
        // Map style index to numFmtId
        if (currentNumFmtIdAttr) {
          const numFmtId = parseInt(currentNumFmtIdAttr, 10);
          if (!isNaN(numFmtId)) {
            styleMap.set(currentStyleIndex, numFmtId);
          }
        }
        currentStyleIndex++;
        currentNumFmtIdAttr = undefined;
      }
    }
  }

  // Build final map: styleIndex -> formatCode
  const result = new Map<number, string>();

  // Map style indices to format codes
  for (const [styleIndex, numFmtId] of styleMap.entries()) {
    // Check if it's a built-in format first
    const builtInCode = getBuiltInFormatCode(numFmtId);
    if (builtInCode) {
      result.set(styleIndex, builtInCode);
    } else {
      // Check custom formats
      const customCode = formatMap.get(numFmtId);
      if (customCode) {
        result.set(styleIndex, customCode);
      }
      // If neither found, don't add to map (will return null in lookup)
    }
  }

  return result;
}

/**
 * Gets format code for a style index from the style format map
 * The map already contains both custom and built-in formats (populated by parseStyles)
 * Returns null if no format code is found for the given style index
 */
export function getFormatCodeForStyle(
  styleIndex: number | undefined,
  styleFormatMap: StyleFormatMap,
): string | null {
  if (styleIndex === undefined) {
    return null;
  }

  // Look up format code in the map (contains both custom and built-in formats)
  // Built-in formats are already included in the map by parseStyles()
  const formatCode = styleFormatMap.get(styleIndex);
  if (formatCode) {
    return formatCode;
  }

  // If not found in map, return null to let the caller use a default format
  // This happens when the style index doesn't have an associated format code
  return null;
}
