import type { Row } from '../types';

/**
 * Maps rows using a transformation function
 */
export async function* mapRows(
  rows: AsyncIterable<Row>,
  fn: (row: Row) => Row | Promise<Row>,
): AsyncIterable<Row> {
  for await (const row of rows) {
    yield await fn(row);
  }
}

/**
 * Filters rows using a predicate function
 */
export async function* filterRows(
  rows: AsyncIterable<Row>,
  fn: (row: Row) => boolean,
): AsyncIterable<Row> {
  for await (const row of rows) {
    if (fn(row)) {
      yield row;
    }
  }
}

/**
 * Limits the number of rows yielded
 */
export async function* limitRows(
  rows: AsyncIterable<Row>,
  limit: number,
): AsyncIterable<Row> {
  let count = 0;
  for await (const row of rows) {
    if (count >= limit) {
      break;
    }
    yield row;
    count++;
  }
}

/**
 * Collects all items from an async iterable into an array
 */
export async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const result: T[] = [];
  for await (const item of iterable) {
    result.push(item);
  }
  return result;
}
