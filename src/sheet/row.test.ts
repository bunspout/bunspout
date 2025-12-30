import { describe, test, expect } from 'bun:test';
import { cell } from './cell';
import { row } from './row';

describe('Row Factory Function', () => {
  test('should create row with cells', () => {
    const result = row([cell('A'), cell(1), cell('B')]);
    expect(result.cells).toHaveLength(3);
    expect(result.cells[0]?.value).toBe('A');
    expect(result.cells[1]?.value).toBe(1);
    expect(result.cells[2]?.value).toBe('B');
  });

  test('should create row with rowIndex option', () => {
    const result = row([cell('Test')], { rowIndex: 5 });
    expect(result.rowIndex).toBe(5);
  });

  test('should create row without options', () => {
    const result = row([cell('Test')]);
    expect(result.rowIndex).toBeUndefined();
  });

  test('should handle empty cells array', () => {
    const result = row([]);
    expect(result.cells).toHaveLength(0);
  });
});
