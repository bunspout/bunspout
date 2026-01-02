/*
 * Validation schemas for XLSX writer inputs
 */
import { z } from 'zod';

/**
 * Validates sheet names according to Excel rules:
 * - Max 31 characters
 * - Cannot contain: :, /, \, ?, *, [, ]
 * - Cannot be empty
 * - Cannot start or end with apostrophe
 */
export const sheetNameSchema = z.string()
  .min(1, 'Sheet name cannot be empty')
  .max(31, 'Sheet name cannot exceed 31 characters')
  .refine(
    (name) => !/[:/\\?*[\]]/.test(name),
    'Sheet name cannot contain :, /, \\, ?, *, [, or ]',
  )
  .refine(
    (name) => !name.startsWith("'") && !name.endsWith("'"),
    'Sheet name cannot start or end with apostrophe',
  );

/**
 * Validates custom property names and values
 * Excel limits:
 * - Property names: max 255 characters, cannot contain certain characters
 * - Property values: max 32767 characters (Excel cell limit)
 */
export const customPropertiesSchema = z.record(
  z.string()
    .min(1, 'Property name cannot be empty')
    .max(255, 'Property name cannot exceed 255 characters')
    .refine(
      (name) => !/[<>:"/\\|?*]/.test(name),
      'Property name contains invalid characters',
    ),
  z.string().max(32767, 'Property value cannot exceed 32767 characters'), // Excel cell limit
);

/**
 * Validates workbook properties
 * Excel limits for various property fields
 */
export const workbookPropertiesSchema = z.object({
  title: z.string().max(255).nullable().optional(),
  subject: z.string().max(255).nullable().optional(),
  application: z.string().max(255).nullable().optional(),
  creator: z.string().max(255).nullable().optional(),
  lastModifiedBy: z.string().max(255).nullable().optional(),
  keywords: z.string().max(255).nullable().optional(),
  description: z.string().max(4000).nullable().optional(), // Excel allows longer descriptions
  category: z.string().max(255).nullable().optional(),
  language: z.string()
    .regex(/^[a-z]{2}(-[A-Z]{2})?$/, 'Language must be ISO 639 format (e.g., "en" or "en-US")')
    .nullable()
    .optional(),
  customProperties: customPropertiesSchema.optional(),
}).strict();
