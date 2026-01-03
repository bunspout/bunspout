/*
 * Runtime-agnostic test framework wrapper
 * Automatically uses Bun's test runner or Node's test runner based on runtime
 */

// Detect runtime
const isBun = typeof Bun !== 'undefined';
const isNode = typeof process !== 'undefined' && process.versions?.node;

// Type definitions that work for both frameworks
export type TestFunction = (name: string, fn: () => void | Promise<void>) => void;
export type DescribeFunction = (name: string, fn: () => void) => void;
export type HookFunction = (fn: () => void | Promise<void>) => void;

export interface ExpectObject {
  toBe: (expected: unknown) => void;
  toEqual: (expected: unknown) => void;
  toContain: (expected: unknown) => void;
  toMatch: (expected: RegExp | string) => void;
  toHaveLength: (expected: number) => void;
  toBeGreaterThan: (expected: number) => void;
  toBeGreaterThanOrEqual: (expected: number) => void;
  toBeLessThan: (expected: number) => void;
  toBeLessThanOrEqual: (expected: number) => void;
  toBeDefined: () => void;
  toBeUndefined: () => void;
  toBeNull: () => void;
  toBeTruthy: () => void;
  toBeFalsy: () => void;
  toBeInstanceOf: (expected: unknown) => void;
  toThrow: (expected?: string | RegExp) => void | Promise<void>;
  rejects: ExpectObject;
  not: ExpectObject;
}

export type ExpectFunction = (actual: unknown) => ExpectObject;

// Runtime-specific imports
let describe: DescribeFunction;
let test: TestFunction;
let expect: ExpectFunction;
let beforeEach: HookFunction;
let afterEach: HookFunction;
let beforeAll: HookFunction;
let afterAll: HookFunction;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let bunTest: any = null;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let nodeTest: any = null;

if (isBun) {
  // Use Bun's test runner
  bunTest = await import('bun:test');
  describe = bunTest.describe;
  test = bunTest.test;
  expect = bunTest.expect;
  beforeEach = bunTest.beforeEach;
  afterEach = bunTest.afterEach;
  beforeAll = bunTest.beforeAll;
  afterAll = bunTest.afterAll;
} else if (isNode) {
  // Use Node's test runner
  nodeTest = await import('node:test');
  const expectModule = await import('expect');

  describe = nodeTest!.describe;
  test = nodeTest!.test;
  beforeEach = nodeTest!.beforeEach;
  afterEach = nodeTest!.afterEach;
  beforeAll = nodeTest!.before;
  afterAll = nodeTest!.after;

  // Use expect package directly - it's Jest-compatible and works standalone
  // Cast to our ExpectFunction type for compatibility (Jest's expect has compatible API)
  expect = (expectModule.default || expectModule.expect) as unknown as ExpectFunction;
} else {
  throw new Error('Unsupported test runtime. This library requires Bun or Node.js.');
}

// Export runtime-agnostic test utilities
export { describe, test, expect, beforeEach, afterEach, beforeAll, afterAll };

// Convenience re-exports
export const it = test;
export const fit = isBun && bunTest
  ? ((name: string, fn: () => void | Promise<void>) => {
    // In CI environments, Bun disables .only, so fall back to regular test
    try {
      return bunTest.test.only(name, fn);
    } catch {
      // Fall back to regular test if .only is disabled (e.g., in CI)
      return bunTest.test(name, fn);
    }
  })
  : isNode && nodeTest && nodeTest.test.only
    ? nodeTest.test.only
    : test; // Fallback if .only is not available
export const xit = isBun && bunTest
  ? bunTest.test.skip
  : ((name: string, fn: () => void | Promise<void> = () => {}) => {
    if (nodeTest?.test.skip) {
      nodeTest.test.skip(name, fn);
    }
  });
