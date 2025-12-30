import { describe, test, expect } from 'bun:test';
import { writeZipEntry, createZipWriter, endZipWriter } from './writer';

describe('ZIP Writer', () => {
  test('should write ZIP entry from AsyncIterable<Uint8Array>', async () => {
    const chunks: Uint8Array[] = [
      new TextEncoder().encode('Hello'),
      new TextEncoder().encode(' '),
      new TextEncoder().encode('World'),
    ];

    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'test.txt', async function* () {
      for (const chunk of chunks) yield chunk;
    }());

    const result = await endZipWriter(zipWriter);
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });

  test('should handle multiple entries', async () => {
    const zipWriter = createZipWriter();

    await writeZipEntry(zipWriter, 'file1.txt', async function* () {
      yield new TextEncoder().encode('Content 1');
    }());

    await writeZipEntry(zipWriter, 'file2.txt', async function* () {
      yield new TextEncoder().encode('Content 2');
    }());

    const result = await endZipWriter(zipWriter);
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });

  test('should handle empty entry', async () => {
    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'empty.txt', async function* () {
      // Empty iterable
    }());

    const result = await endZipWriter(zipWriter);
    expect(result).toBeInstanceOf(Buffer);
  });

  test('should handle large chunks', async () => {
    const largeData = new Uint8Array(10000).fill(65); // 10KB of 'A'
    const zipWriter = createZipWriter();

    await writeZipEntry(zipWriter, 'large.txt', async function* () {
      yield largeData;
    }());

    const result = await endZipWriter(zipWriter);
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });
});
