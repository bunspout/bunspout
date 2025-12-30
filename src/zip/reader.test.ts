import { describe, test, expect } from 'bun:test';
import { readZipEntry, openZip } from './reader';
import { createZipWriter, writeZipEntry, endZipWriter } from './writer';

describe('ZIP Reader', () => {
  test('should read ZIP entry as AsyncIterable<Uint8Array>', async () => {
    // Create a test ZIP file
    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'test.txt', async function* () {
      yield new TextEncoder().encode('Hello World');
    }());
    const zipBuffer = await endZipWriter(zipWriter);

    // Read it back
    const zipFile = await openZip(zipBuffer);
    const entry = zipFile.entries.find((e) => e.fileName === 'test.txt');
    expect(entry).toBeDefined();

    if (entry) {
      const chunks: Uint8Array[] = [];
      for await (const chunk of readZipEntry(entry, zipFile.zipFile)) {
        chunks.push(chunk);
      }

      const result = new TextDecoder().decode(
        new Uint8Array(
          chunks.reduce((acc, chunk) => {
            const combined = new Uint8Array(acc.length + chunk.length);
            combined.set(acc);
            combined.set(chunk, acc.length);
            return combined;
          }, new Uint8Array(0)),
        ),
      );

      expect(result).toBe('Hello World');
    }
  });

  test('should handle multiple entries', async () => {
    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'file1.txt', async function* () {
      yield new TextEncoder().encode('Content 1');
    }());
    await writeZipEntry(zipWriter, 'file2.txt', async function* () {
      yield new TextEncoder().encode('Content 2');
    }());
    const zipBuffer = await endZipWriter(zipWriter);

    const zipFile = await openZip(zipBuffer);
    expect(zipFile.entries.length).toBe(2);

    const file1 = zipFile.entries.find((e) => e.fileName === 'file1.txt');
    expect(file1).toBeDefined();

    if (file1) {
      const chunks: Uint8Array[] = [];
      for await (const chunk of readZipEntry(file1, zipFile.zipFile)) {
        chunks.push(chunk);
      }
      const result = new TextDecoder().decode(
        Buffer.concat(chunks.map((c) => Buffer.from(c))),
      );
      expect(result).toBe('Content 1');
    }
  });

  test('should handle empty entry', async () => {
    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'empty.txt', async function* () {
      // Empty
    }());
    const zipBuffer = await endZipWriter(zipWriter);

    const zipFile = await openZip(zipBuffer);
    const entry = zipFile.entries.find((e) => e.fileName === 'empty.txt');
    expect(entry).toBeDefined();

    if (entry) {
      const chunks: Uint8Array[] = [];
      for await (const chunk of readZipEntry(entry, zipFile.zipFile)) {
        chunks.push(chunk);
      }
      expect(chunks.length).toBe(0);
    }
  });

  test('should handle large entries', async () => {
    const largeData = new Uint8Array(10000).fill(65); // 10KB of 'A'
    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'large.txt', async function* () {
      yield largeData;
    }());
    const zipBuffer = await endZipWriter(zipWriter);

    const zipFile = await openZip(zipBuffer);
    const entry = zipFile.entries.find((e) => e.fileName === 'large.txt');
    expect(entry).toBeDefined();

    if (entry) {
      const chunks: Uint8Array[] = [];
      for await (const chunk of readZipEntry(entry, zipFile.zipFile)) {
        chunks.push(chunk);
      }
      const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
      expect(totalLength).toBe(10000);
    }
  });
});
