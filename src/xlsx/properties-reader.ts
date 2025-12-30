import type * as yauzl from 'yauzl';
import { readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import type { WorkbookProperties } from './types';

/**
 * Parses core.xml to extract workbook properties
 */
export async function parseCoreProperties(
  zipEntry: ZipEntry,
  zipFile: yauzl.ZipFile,
): Promise<WorkbookProperties> {
  const properties: WorkbookProperties = {};

  let currentElement: string | null = null;
  let currentText = '';

  for await (const event of parseXmlEvents(readZipEntry(zipEntry, zipFile))) {
    if (event.type === 'startElement') {
      currentElement = event.name || null;
      currentText = '';
    } else if (event.type === 'text' && currentElement) {
      currentText += event.text || '';
    } else if (event.type === 'endElement' && currentElement) {
      const text = currentText.trim();
      if (text) {
        // Handle both prefixed (dc:title, cp:keywords) and unprefixed (title, keywords) element names
        const localName = currentElement.includes(':') ? currentElement.split(':')[1] : currentElement;
        switch (localName) {
          case 'title':
            properties.title = text;
            break;
          case 'subject':
            properties.subject = text;
            break;
          case 'creator':
            properties.creator = text;
            break;
          case 'lastModifiedBy':
            properties.lastModifiedBy = text;
            break;
          case 'keywords':
            properties.keywords = text;
            break;
          case 'description':
            properties.description = text;
            break;
          case 'category':
            properties.category = text;
            break;
          case 'language':
            properties.language = text;
            break;
        }
      }
      currentElement = null;
      currentText = '';
    }
  }

  return properties;
}

/**
 * Parses custom.xml to extract custom properties
 */
export async function parseCustomProperties(
  zipEntry: ZipEntry,
  zipFile: yauzl.ZipFile,
): Promise<Record<string, string>> {
  const customProperties: Record<string, string> = {};

  let inProperty = false;
  let propertyName: string | null = null;
  let propertyValue: string | null = null;
  let inValue = false;
  let currentText = '';

  for await (const event of parseXmlEvents(readZipEntry(zipEntry, zipFile))) {
    if (event.type === 'startElement') {
      if (event.name === 'property') {
        inProperty = true;
        propertyName = event.attributes?.name || null;
        propertyValue = null;
      } else if (event.name === 'vt:lpwstr' && inProperty) {
        inValue = true;
        currentText = '';
      }
    } else if (event.type === 'text' && inValue) {
      currentText += event.text || '';
    } else if (event.type === 'endElement') {
      if (event.name === 'vt:lpwstr' && inValue) {
        propertyValue = currentText.trim();
        inValue = false;
        currentText = '';
      } else if (event.name === 'property' && inProperty && propertyName && propertyValue) {
        customProperties[propertyName] = propertyValue;
        inProperty = false;
        propertyName = null;
        propertyValue = null;
      }
    }
  }

  return customProperties;
}
