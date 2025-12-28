import type * as yauzl from 'yauzl';
import { readZipEntry, type ZipEntry } from '@zip/reader';
import { parseXmlEvents } from '@xml/parser';
import { CachingStrategyFactory, type SharedStringsCachingStrategy } from './shared-strings-caching';

/**
 * Parses shared strings XML and returns a caching strategy
 */
export async function parseSharedStrings(
  zipEntry: ZipEntry,
  zipFile: yauzl.ZipFile,
): Promise<SharedStringsCachingStrategy> {
  // First pass: try to get uniqueCount from <sst> element attributes
  let uniqueCount: number | null = null;
  let currentIndex = 0;
  let currentString = '';
  let inStringItem = false;
  let inText = false;
  let strategy: SharedStringsCachingStrategy | null = null;

  for await (const event of parseXmlEvents(readZipEntry(zipEntry, zipFile))) {
    if (event.type === 'startElement') {
      if (event.name === 'sst') {
        // Try to get uniqueCount from attributes
        const uniqueCountAttr = event.attributes?.['uniqueCount'];
        if (uniqueCountAttr) {
          const parsed = parseInt(uniqueCountAttr, 10);
          if (!isNaN(parsed)) {
            uniqueCount = parsed;
            // Create strategy now that we know the count
            strategy = CachingStrategyFactory.createBestCachingStrategy(uniqueCount);
          }
        }
      } else if (event.name === 'si') {
        inStringItem = true;
        currentString = '';
      } else if (event.name === 't' && inStringItem) {
        inText = true;
      }
    } else if (event.type === 'endElement') {
      if (event.name === 'si' && inStringItem) {
        // If strategy not created yet (uniqueCount was not in XML), create it now
        if (!strategy) {
          strategy = CachingStrategyFactory.createBestCachingStrategy(null);
        }
        // Add string to strategy
        await strategy.addString(currentIndex, currentString);
        currentIndex++;
        inStringItem = false;
        currentString = '';
      } else if (event.name === 't' && inText) {
        inText = false;
      }
    } else if (event.type === 'text' && inText && inStringItem) {
      currentString += event.text || '';
    }
  }

  // If no strategy was created (empty file), create one
  if (!strategy) {
    strategy = CachingStrategyFactory.createBestCachingStrategy(0);
  }

  return strategy;
}

