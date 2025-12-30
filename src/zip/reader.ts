import { Readable } from 'stream';
import * as yauzl from 'yauzl';
import { nodeStreamToBytes } from '../adapters/node';

export type ZipEntry = {
  fileName: string;
  entry: yauzl.Entry;
};

export type ZipFile = {
  zipFile: yauzl.ZipFile;
  entries: ZipEntry[];
};

/**
 * Opens a ZIP file from a buffer
 */
export function openZip(buffer: Buffer): Promise<ZipFile> {
  return new Promise((resolve, reject) => {
    yauzl.fromBuffer(buffer, { lazyEntries: true }, (err, zipfile) => {
      if (err) {
        reject(err);
        return;
      }
      if (!zipfile) {
        reject(new Error('Failed to open ZIP file'));
        return;
      }

      const entries: ZipEntry[] = [];

      zipfile.on('entry', (entry: yauzl.Entry) => {
        entries.push({
          fileName: entry.fileName,
          entry,
        });
        zipfile.readEntry();
      });

      zipfile.on('end', () => {
        resolve({
          zipFile: zipfile,
          entries,
        });
      });

      zipfile.on('error', reject);
      zipfile.readEntry();
    });
  });
}

/**
 * Reads a ZIP entry as AsyncIterable<Uint8Array>
 */
export async function* readZipEntry(
  zipEntry: ZipEntry,
  zipfile: yauzl.ZipFile,
): AsyncIterable<Uint8Array> {
  const stream = await new Promise<Readable>((resolve, reject) => {
    zipfile.openReadStream(zipEntry.entry, (err, stream) => {
      if (err) {
        reject(err);
        return;
      }
      if (!stream) {
        reject(new Error('Failed to open read stream'));
        return;
      }
      resolve(stream);
    });
  });

  yield* nodeStreamToBytes(stream);
}
