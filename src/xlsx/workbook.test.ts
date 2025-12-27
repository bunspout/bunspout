/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect } from 'bun:test';
import { Workbook } from './workbook';

describe('Workbook and Sheet Classes', () => {
  test('should create Workbook from ZIP file', () => {
    // Mock ZIP file structure
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/workbook.xml', entry: {} as any },
      ],
    };

    const workbook = new Workbook(mockZipFile, [
      { name: 'Sheet1', entry: mockZipFile.entries[0]! },
    ]);

    expect(workbook).toBeInstanceOf(Workbook);
  });

  test('should get sheet by name', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/worksheets/sheet2.xml', entry: {} as any },
      ],
    };

    const sheets = [
      { name: 'Data', entry: mockZipFile.entries[0]! },
      { name: 'Summary', entry: mockZipFile.entries[1]! },
    ];

    const workbook = new Workbook(mockZipFile, sheets);

    const sheet = workbook.sheet('Data');
    expect(sheet.name).toBe('Data');
  });

  test('should get sheet by index', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/worksheets/sheet2.xml', entry: {} as any },
      ],
    };

    const sheets = [
      { name: 'Data', entry: mockZipFile.entries[0]! },
      { name: 'Summary', entry: mockZipFile.entries[1]! },
    ];

    const workbook = new Workbook(mockZipFile, sheets);

    const sheet = workbook.sheet(0);
    expect(sheet.name).toBe('Data');

    const sheet2 = workbook.sheet(1);
    expect(sheet2.name).toBe('Summary');
  });

  test('should get all sheets', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [
        { fileName: 'xl/worksheets/sheet1.xml', entry: {} as any },
        { fileName: 'xl/worksheets/sheet2.xml', entry: {} as any },
      ],
    };

    const sheets = [
      { name: 'Data', entry: mockZipFile.entries[0]! },
      { name: 'Summary', entry: mockZipFile.entries[1]! },
    ];

    const workbook = new Workbook(mockZipFile, sheets);

    const allSheets = workbook.sheets();
    expect(allSheets).toHaveLength(2);
    expect(allSheets[0]?.name).toBe('Data');
    expect(allSheets[1]?.name).toBe('Summary');
  });

  test('should throw error for missing sheet by name', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [{ fileName: 'xl/worksheets/sheet1.xml', entry: {} as any }],
    };

    const workbook = new Workbook(mockZipFile, [
      { name: 'Data', entry: mockZipFile.entries[0]! },
    ]);

    expect(() => workbook.sheet('NonExistent')).toThrow();
  });

  test('should throw error for missing sheet by index', () => {
    const mockZipFile = {
      zipFile: {} as any,
      entries: [{ fileName: 'xl/worksheets/sheet1.xml', entry: {} as any }],
    };

    const workbook = new Workbook(mockZipFile, [
      { name: 'Data', entry: mockZipFile.entries[0]! },
    ]);

    expect(() => workbook.sheet(99)).toThrow();
  });
});

