/**
 * XML utility functions
 */

/**
 * Removes invalid XML control characters (keeps tab, newline, carriage return)
 * According to XML 1.0 spec, only 0x09 (tab), 0x0A (newline), 0x0D (CR) are allowed
 * Control chars 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F are invalid
 */
export function removeInvalidControlChars(text: string): string {
  // eslint-disable-next-line no-control-regex -- It's exactly what we want to match in this case
  return text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
}

/**
 * Escapes XML special characters and removes invalid control characters
 */
export function escapeXml(text: string): string {
  // First remove invalid control characters
  const cleaned = removeInvalidControlChars(text);
  // Then escape XML special characters
  return cleaned
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
