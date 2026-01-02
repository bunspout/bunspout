/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect } from 'bun:test';
import { sheetNameSchema, workbookPropertiesSchema, customPropertiesSchema } from './validation';

describe('Validation Schemas', () => {
  describe('sheetNameSchema', () => {
    test('should accept valid sheet names', () => {
      expect(() => sheetNameSchema.parse('Sheet1')).not.toThrow();
      expect(() => sheetNameSchema.parse('Data')).not.toThrow();
      expect(() => sheetNameSchema.parse('My Sheet')).not.toThrow();
      expect(() => sheetNameSchema.parse('A'.repeat(31))).not.toThrow(); // Max length
    });

    test('should reject empty sheet names', () => {
      expect(() => sheetNameSchema.parse('')).toThrow('cannot be empty');
    });

    test('should reject sheet names exceeding 31 characters', () => {
      expect(() => sheetNameSchema.parse('A'.repeat(32))).toThrow('cannot exceed 31 characters');
    });

    test('should reject sheet names with invalid characters', () => {
      expect(() => sheetNameSchema.parse('Sheet:1')).toThrow('cannot contain');
      expect(() => sheetNameSchema.parse('Sheet/1')).toThrow('cannot contain');
      expect(() => sheetNameSchema.parse('Sheet\\1')).toThrow('cannot contain');
      expect(() => sheetNameSchema.parse('Sheet?1')).toThrow('cannot contain');
      expect(() => sheetNameSchema.parse('Sheet*1')).toThrow('cannot contain');
      expect(() => sheetNameSchema.parse('Sheet[1]')).toThrow('cannot contain');
    });

    test('should reject sheet names starting with apostrophe', () => {
      expect(() => sheetNameSchema.parse("'Sheet")).toThrow('cannot start or end with apostrophe');
    });

    test('should reject sheet names ending with apostrophe', () => {
      expect(() => sheetNameSchema.parse("Sheet'")).toThrow('cannot start or end with apostrophe');
    });

    test('should reject sheet names with apostrophes on both ends', () => {
      expect(() => sheetNameSchema.parse("'Sheet'")).toThrow('cannot start or end with apostrophe');
    });

    test('should accept sheet names with apostrophes in the middle', () => {
      expect(() => sheetNameSchema.parse("Sheet'Name")).not.toThrow();
    });
  });

  describe('workbookPropertiesSchema', () => {
    test('should accept valid workbook properties', () => {
      expect(() => workbookPropertiesSchema.parse({
        title: 'My Workbook',
        creator: 'Author',
        subject: 'Testing',
      })).not.toThrow();
    });

    test('should accept null values', () => {
      expect(() => workbookPropertiesSchema.parse({
        title: null,
        creator: null,
      })).not.toThrow();
    });

    test('should accept empty object', () => {
      expect(() => workbookPropertiesSchema.parse({})).not.toThrow();
    });

    test('should reject properties exceeding max length', () => {
      expect(() => workbookPropertiesSchema.parse({
        title: 'A'.repeat(256), // Max is 255
      })).toThrow();
    });

    test('should accept description up to 4000 characters', () => {
      expect(() => workbookPropertiesSchema.parse({
        description: 'A'.repeat(4000),
      })).not.toThrow();
    });

    test('should reject description exceeding 4000 characters', () => {
      expect(() => workbookPropertiesSchema.parse({
        description: 'A'.repeat(4001),
      })).toThrow();
    });

    test('should accept valid language codes', () => {
      expect(() => workbookPropertiesSchema.parse({
        language: 'en',
      })).not.toThrow();
      expect(() => workbookPropertiesSchema.parse({
        language: 'en-US',
      })).not.toThrow();
      expect(() => workbookPropertiesSchema.parse({
        language: 'fr',
      })).not.toThrow();
      expect(() => workbookPropertiesSchema.parse({
        language: 'zh-CN',
      })).not.toThrow();
    });

    test('should reject invalid language codes', () => {
      expect(() => workbookPropertiesSchema.parse({
        language: 'invalid',
      })).toThrow('Language must be ISO 639 format');
      expect(() => workbookPropertiesSchema.parse({
        language: 'EN',
      })).toThrow('Language must be ISO 639 format');
      expect(() => workbookPropertiesSchema.parse({
        language: 'en-us', // Should be en-US
      })).toThrow('Language must be ISO 639 format');
    });

    test('should validate custom properties when present', () => {
      expect(() => workbookPropertiesSchema.parse({
        customProperties: {
          'Custom1': 'Value1',
          'Custom2': 'Value2',
        },
      })).not.toThrow();
    });

    test('should reject unknown properties', () => {
      expect(() => workbookPropertiesSchema.parse({
        title: 'Test',
        unknownField: 'value',
      } as any)).toThrow();
    });
  });

  describe('customPropertiesSchema', () => {
    test('should accept valid custom properties', () => {
      expect(() => customPropertiesSchema.parse({
        'Property1': 'Value1',
        'Property2': 'Value2',
      })).not.toThrow();
    });

    test('should reject empty property names', () => {
      expect(() => customPropertiesSchema.parse({
        '': 'Value',
      })).toThrow('Property name cannot be empty');
    });

    test('should reject property names exceeding 255 characters', () => {
      expect(() => customPropertiesSchema.parse({
        ['A'.repeat(256)]: 'Value',
      })).toThrow('Property name cannot exceed 255 characters');
    });

    test('should reject property names with invalid characters', () => {
      expect(() => customPropertiesSchema.parse({
        'Property<Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property>Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property:Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property"Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property/Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property\\Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property|Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property?Name': 'Value',
      })).toThrow('Property name contains invalid characters');
      expect(() => customPropertiesSchema.parse({
        'Property*Name': 'Value',
      })).toThrow('Property name contains invalid characters');
    });

    test('should accept property values up to 32767 characters', () => {
      expect(() => customPropertiesSchema.parse({
        'Property': 'A'.repeat(32767),
      })).not.toThrow();
    });

    test('should reject property values exceeding 32767 characters', () => {
      expect(() => customPropertiesSchema.parse({
        'Property': 'A'.repeat(32768),
      })).toThrow('Property value cannot exceed 32767 characters');
    });
  });
});
