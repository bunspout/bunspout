import * as yazl from 'yazl';

/**
 * Creates a new ZIP writer
 */
export function createZipWriter(): yazl.ZipFile {
  return new yazl.ZipFile();
}

/**
 * Writes a ZIP entry from an AsyncIterable<Uint8Array>
 */
export async function writeZipEntry(
  zipFile: yazl.ZipFile,
  name: string,
  bytes: AsyncIterable<Uint8Array>,
): Promise<void> {
  // Collect all chunks into a single buffer
  const chunks: Uint8Array[] = [];
  for await (const chunk of bytes) {
    chunks.push(chunk);
  }

  // Combine chunks into a single buffer
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const buffer = Buffer.allocUnsafe(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    buffer.set(chunk, offset);
    offset += chunk.length;
  }

  // Add buffer to ZIP
  zipFile.addBuffer(buffer, name);
}

/**
 * Ends the ZIP file and returns the buffer
 */
export function endZipWriter(zipFile: yazl.ZipFile): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    zipFile.outputStream.on('data', (chunk: Buffer) => {
      chunks.push(chunk);
    });
    zipFile.outputStream.on('end', () => {
      resolve(Buffer.concat(chunks));
    });
    zipFile.outputStream.on('error', reject);
    zipFile.end();
  });
}
