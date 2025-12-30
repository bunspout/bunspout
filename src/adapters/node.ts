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

/**
 * Writes a buffer to a file (Node.js implementation)
 */
export async function writeFile(filePath: string, buffer: Uint8Array | Buffer): Promise<void> {
  const { writeFile: fsWriteFile } = await import('fs/promises');
  await fsWriteFile(filePath, buffer);
}

/**
 * Reads a file to a Buffer (Node.js implementation)
 */
export async function readFile(filePath: string): Promise<Buffer> {
  const { readFile: fsReadFile } = await import('fs/promises');
  const buffer = await fsReadFile(filePath);
  return Buffer.from(buffer);
}
