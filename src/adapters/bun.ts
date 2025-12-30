import { writeSheetXml } from '@xml/writer';
import type { Row } from '../types';

/**
 * Converts AsyncIterable<string> to ReadableStream for Bun
 */
export function stringToReadableStream(
  strings: AsyncIterable<string>,
): ReadableStream<Uint8Array> {
  return new ReadableStream({
    async start(controller) {
      try {
        for await (const str of strings) {
          const bytes = new TextEncoder().encode(str);
          controller.enqueue(bytes);
        }
        controller.close();
      } catch (error) {
        controller.error(error);
      }
    },
  });
}

/**
 * Creates a Bun Response from rows
 */
export function rowsToResponse(rows: AsyncIterable<Row>): Response {
  const xmlStream = writeSheetXml(rows);
  const bytesStream = stringToReadableStream(xmlStream);
  return new Response(bytesStream, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  });
}

/**
 * Converts Blob to AsyncIterable<Uint8Array>
 */
export async function* blobToBytes(blob: Blob): AsyncIterable<Uint8Array> {
  const buffer = await blob.arrayBuffer();
  yield new Uint8Array(buffer);
}

/**
 * Converts ReadableStream to AsyncIterable<Uint8Array>
 */
export async function* readableStreamToBytes(
  stream: ReadableStream<Uint8Array>,
): AsyncIterable<Uint8Array> {
  const reader = stream.getReader();
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (value) {
        yield value;
      }
    }
  } finally {
    reader.releaseLock();
  }
}

/**
 * Writes a buffer to a file (Bun implementation)
 */
export async function writeFile(filePath: string, buffer: Uint8Array | Buffer): Promise<void> {
  await Bun.write(filePath, buffer);
}

/**
 * Reads a file to a Buffer (Bun implementation)
 */
export async function readFile(filePath: string): Promise<Buffer> {
  const file = Bun.file(filePath);
  if (!(await file.exists())) {
    throw new Error(`File not found: ${filePath}`);
  }
  const arrayBuffer = await file.arrayBuffer();
  return Buffer.from(arrayBuffer);
}
