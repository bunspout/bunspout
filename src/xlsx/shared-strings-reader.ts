import type * as yauzl from 'yauzl';
import { readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';

/**
 * Parses shared strings XML and returns an array of strings
 */
export async function parseSharedStrings(
  zipEntry: ZipEntry,
  zipFile: yauzl.ZipFile,
): Promise<string[]> {
  // Stream XML directly from ZIP to parser (no accumulation)
  const strings: string[] = [];
  let currentString = '';
  let inStringItem = false;
  let inText = false;

  for await (const event of parseXmlEvents(readZipEntry(zipEntry, zipFile))) {
    if (event.type === 'startElement') {
      if (event.name === 'si') {
        inStringItem = true;
        currentString = '';
      } else if (event.name === 't' && inStringItem) {
        inText = true;
      }
    } else if (event.type === 'endElement') {
      if (event.name === 'si' && inStringItem) {
        strings.push(currentString);
        inStringItem = false;
        currentString = '';
      } else if (event.name === 't' && inText) {
        inText = false;
      }
    } else if (event.type === 'text' && inText && inStringItem) {
      currentString += event.text || '';
    }
  }

  return strings;
}

