/*
 * Runtime detection and adapter exports
 */

const isBun = typeof Bun !== 'undefined';
const isNode = typeof process !== 'undefined' && process.versions?.node;

/**
 * Writes a buffer to a file using the appropriate runtime adapter
 */
export async function writeFile(filePath: string, buffer: Uint8Array | Buffer): Promise<void> {
  if (isBun) {
    const { writeFile } = await import('./bun');
    return writeFile(filePath, buffer);
  } else if (isNode) {
    const { writeFile } = await import('./node');
    return writeFile(filePath, buffer);
  } else {
    throw new Error('Unsupported runtime. This library requires Bun or Node.js.');
  }
}

/**
 * Reads a file to a Buffer using the appropriate runtime adapter
 */
export async function readFile(filePath: string): Promise<Buffer> {
  if (isBun) {
    const { readFile } = await import('./bun');
    return readFile(filePath);
  } else if (isNode) {
    const { readFile } = await import('./node');
    return readFile(filePath);
  } else {
    throw new Error('Unsupported runtime. This library requires Bun or Node.js.');
  }
}

// Re-export common utilities
export * from './common';
