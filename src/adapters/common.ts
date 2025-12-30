/**
 * Common adapter utilities for converting between different data formats
 */

/**
 * Converts a string to AsyncIterable<Uint8Array>
 */
export async function* stringToBytes(str: string): AsyncIterable<Uint8Array> {
  yield new TextEncoder().encode(str);
}

/**
 * Converts AsyncIterable<Uint8Array> to a string
 */
export async function bytesToString(
  bytes: AsyncIterable<Uint8Array>,
): Promise<string> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of bytes) {
    chunks.push(chunk);
  }
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return new TextDecoder().decode(result);
}

/**
 * Converts AsyncIterable<Uint8Array> to a single Uint8Array
 */
export async function bytesToUint8Array(
  bytes: AsyncIterable<Uint8Array>,
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of bytes) {
    chunks.push(chunk);
  }
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

/**
 * Converts Uint8Array to AsyncIterable<Uint8Array>
 */
export async function* uint8ArrayToBytes(
  data: Uint8Array,
): AsyncIterable<Uint8Array> {
  yield data;
}
