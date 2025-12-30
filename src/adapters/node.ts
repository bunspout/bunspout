import { Readable } from 'stream';
import { writeSheetXml } from '@xml/writer';
import type { Row } from '../types';

/**
 * Converts AsyncIterable<string> to Node.js Readable stream
 */
export function stringToNodeStream(
  strings: AsyncIterable<string>,
): Readable {
  return Readable.from(
    (async function* () {
      for await (const str of strings) {
        yield str;
      }
    })(),
  );
}

/**
 * Converts rows to Node.js Readable stream
 */
export function rowsToNodeStream(rows: AsyncIterable<Row>): Readable {
  const xmlStream = writeSheetXml(rows);
  return stringToNodeStream(xmlStream);
}

/**
 * Converts Node.js Readable stream to AsyncIterable<Uint8Array>
 */
export function nodeStreamToBytes(
  stream: Readable,
): AsyncIterable<Uint8Array> {
  return (async function* () {
    for await (const chunk of stream) {
      yield new Uint8Array(chunk);
    }
  })();
}

/**
 * Helper to write rows to a file using Node.js streams
 */
export async function writeRowsToFile(
  rows: AsyncIterable<Row>,
  filePath: string,
): Promise<void> {
  const { createWriteStream } = await import('fs');
  const { pipeline } = await import('stream/promises');
  const stream = rowsToNodeStream(rows);
  const writeStream = createWriteStream(filePath);
  await pipeline(stream, writeStream);
}
