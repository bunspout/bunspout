// noinspection HtmlDeprecatedTag,XmlDeprecatedElement

/*
 * Style registry and styles.xml generator
 */
import { z } from 'zod';
import { escapeXml } from '@utils/xml';
import type { Style, FontStyle } from '../types';

// Zod schema for font style validation
const fontStyleSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  fontSize: z.number().int().positive().max(409).optional(), // Excel max is 409
  fontColor: z.string().regex(/^[0-9A-Fa-f]{8}$/, 'Font color must be 8 hex characters (ARGB format)').optional(),
  fontName: z.string().min(1, 'Font name cannot be empty').max(31, 'Font name cannot exceed 31 characters').optional(),
}).strict();

// Default font style values
const DEFAULT_FONT: Required<FontStyle> = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontSize: 11,
  fontColor: 'FF000000', // Black in ARGB
  fontName: 'Arial',
};

/**
 * Normalizes and validates a font style by filling in defaults
 * @throws {z.ZodError} If font style validation fails
 */
function normalizeFontStyle(font?: FontStyle): Required<FontStyle> {
  if (!font) {
    return DEFAULT_FONT;
  }

  const validated = fontStyleSchema.parse(font);

  return {
    bold: validated.bold ?? DEFAULT_FONT.bold,
    italic: validated.italic ?? DEFAULT_FONT.italic,
    underline: validated.underline ?? DEFAULT_FONT.underline,
    strikethrough: validated.strikethrough ?? DEFAULT_FONT.strikethrough,
    fontSize: validated.fontSize ?? DEFAULT_FONT.fontSize,
    fontColor: validated.fontColor ?? DEFAULT_FONT.fontColor,
    fontName: validated.fontName ?? DEFAULT_FONT.fontName,
  };
}

/**
 * Checks if a font style should be applied (differs from default)
 */
function shouldApplyFont(font: Required<FontStyle>): boolean {
  return (
    font.bold !== DEFAULT_FONT.bold ||
    font.italic !== DEFAULT_FONT.italic ||
    font.underline !== DEFAULT_FONT.underline ||
    font.strikethrough !== DEFAULT_FONT.strikethrough ||
    font.fontSize !== DEFAULT_FONT.fontSize ||
    font.fontColor !== DEFAULT_FONT.fontColor ||
    font.fontName !== DEFAULT_FONT.fontName
  );
}

/**
 * Internal representation of a normalized style
 */
interface NormalizedStyle {
  font: Required<FontStyle>;
}

/**
 * Deep freezes an object to ensure structural immutability
 * This prevents accidental mutations of stored styles
 */
function deepFreeze<T>(obj: T): Readonly<T> {
  const propNames = Object.getOwnPropertyNames(obj);

  // Freeze properties before freezing self
  for (const name of propNames) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const value = (obj as any)[name];
    if (value && typeof value === 'object') {
      deepFreeze(value);
    }
  }

  return Object.freeze(obj) as Readonly<T>;
}

/**
 * Style registry that tracks and deduplicates styles.
 * Styles are deep-frozen to prevent mutations.
 * The registry generates the styles.xml content for XLSX files.
 */
export class StyleRegistry {
  private styles: NormalizedStyle[] = [];
  private indexMap: Map<string, number> = new Map();
  private cachedFonts: Required<FontStyle>[] | null = null;

  /**
   * Adds a style to the registry if needed and returns the relevant cellXfs index.
   * @param style - The style to add
   * @returns The cellXfs index to use in XML (1-based, since 0 is reserved for default style)
   * @throws {z.ZodError} If style validation fails
   */
  addStyle(style: Style): number {
    const normalized: NormalizedStyle = {
      font: normalizeFontStyle(style.font),
    };

    const key = this.getStyleKey(normalized);

    if (this.indexMap.has(key)) {
      const registryIndex = this.indexMap.get(key)!;
      // Return cellXfs index (registry index + 1, since 0 is reserved for default)
      return registryIndex + 1;
    }

    // Freeze to prevent mutations that could break deduplication or ordering
    const frozen = deepFreeze(normalized);

    // Add new style. Keep track of index
    const registryIndex = this.styles.length;
    this.styles.push(frozen);
    this.indexMap.set(key, registryIndex);
    // Invalidate font cache when new style is added
    this.cachedFonts = null;
    // Return cellXfs index (registry index + 1, since 0 is reserved for default)
    return registryIndex + 1;
  }

  /**
   * Gets the cellXfs index for a style (must have been added first).
   * @param style - The style to look up
   * @returns The cellXfs index to use in XML (1-based, since 0 is reserved for default style)
   * @throws {Error} If style is not found in registry
   */
  getCellXfIndex(style: Style): number {
    const normalized: NormalizedStyle = {
      font: normalizeFontStyle(style.font),
    };
    const key = this.getStyleKey(normalized);
    const registryIndex = this.indexMap.get(key);
    if (registryIndex === undefined) {
      const fontInfo = style.font
        ? `font: ${style.font.fontName ?? 'default'}, size: ${style.font.fontSize ?? 'default'}, bold: ${style.font.bold ?? false}`
        : 'empty style';
      throw new Error(`Style not found in style registry. Style must be added via addStyle() before calling getCellXfIndex(). Style: ${fontInfo}`);
    }
    // Return cellXfs index (registry index + 1, since 0 is reserved for default)
    return registryIndex + 1;
  }

  /**
   * Gets all registered styles in order
   * @returns Read-only array of normalized styles
   */
  getStyles(): readonly NormalizedStyle[] {
    return this.styles;
  }

  /**
   * Gets the count of unique styles registered
   */
  getCount(): number {
    return this.styles.length;
  }

  /**
   * Generates a unique key for style deduplication
   */
  private getStyleKey(style: NormalizedStyle): string {
    const f = style.font;
    return `font:${f.bold}:${f.italic}:${f.underline}:${f.strikethrough}:${f.fontSize}:${f.fontColor}:${f.fontName}`;
  }

  /**
   * Generates the fonts section of styles.xml
   * Returns fonts in order with deduplication
   * Results are cached to avoid recomputation
   */
  private generateFontsSection(): Required<FontStyle>[] {
    if (this.cachedFonts !== null) {
      return this.cachedFonts;
    }

    // Always include default font (index 0)
    const fonts: Required<FontStyle>[] = [DEFAULT_FONT];
    const fontKeys = new Set<string>();
    fontKeys.add(this.getFontKey(DEFAULT_FONT));

    // Add fonts from registered styles (deduplicated)
    for (const style of this.styles) {
      if (shouldApplyFont(style.font)) {
        const fontKey = this.getFontKey(style.font);
        if (!fontKeys.has(fontKey)) {
          fonts.push(style.font);
          fontKeys.add(fontKey);
        }
      }
    }

    this.cachedFonts = fonts;
    return fonts;
  }

  /**
   * Generates the fonts XML section
   */
  private generateFontsXml(): string {
    const fonts = this.generateFontsSection();

    const fontElements = fonts.map((font) => {
      const elements: string[] = [];
      elements.push(`      <sz val="${font.fontSize}"/>`);
      elements.push(`      <color rgb="${font.fontColor}"/>`);
      elements.push(`      <name val="${escapeXml(font.fontName)}"/>`);

      if (font.bold) {
        elements.push('      <b/>');
      }
      if (font.italic) {
        elements.push('      <i/>');
      }
      if (font.underline) {
        elements.push('      <u/>');
      }
      if (font.strikethrough) {
        elements.push('      <strike/>');
      }

      return `    <font>\n${elements.join('\n')}\n    </font>`;
    });

    return `  <fonts count="${fonts.length}">\n${fontElements.join('\n')}\n  </fonts>`;
  }

  /**
   * Generates the cellXfs section of styles.xml
   * Maps style indices to font indices
   */
  private generateCellXfsSection(): string {
    // Get fonts in the same order as generateFontsSection
    const fonts = this.generateFontsSection();

    // Build a map of font key to font index
    const fontToIndex = new Map<string, number>();
    fonts.forEach((font, index) => {
      fontToIndex.set(this.getFontKey(font), index);
    });

    // Default style (index 0) uses default font (index 0)
    const xfElements: string[] = ['    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'];

    // Map each registered style to its font index
    for (const style of this.styles) {
      let fontId = 0; // Default to default font
      if (shouldApplyFont(style.font)) {
        const fontKey = this.getFontKey(style.font);
        fontId = fontToIndex.get(fontKey) ?? 0;
      }
      xfElements.push(`    <xf numFmtId="0" fontId="${fontId}" fillId="0" borderId="0"/>`);
    }

    return `  <cellXfs count="${xfElements.length}">\n${xfElements.join('\n')}\n  </cellXfs>`;
  }

  /**
   * Generates a unique key for a font style
   */
  private getFontKey(font: Required<FontStyle>): string {
    return `${font.bold}:${font.italic}:${font.underline}:${font.strikethrough}:${font.fontSize}:${font.fontColor}:${font.fontName}`;
  }

  /**
   * Generates the complete styles.xml content for XLSX files.
   * Includes fonts, fills, borders, and cellXfs sections.
   * @returns XML string for styles.xml
   */
  generateXml(): string {
    // If no styles registered, return minimal styles.xml
    if (this.styles.length === 0) {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color rgb="FF000000"/>
      <name val="Arial"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;
    }

    const fontsSection = this.generateFontsXml();
    const cellXfsSection = this.generateCellXfsSection();

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${fontsSection}
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
${cellXfsSection}
</styleSheet>`;
  }
}
