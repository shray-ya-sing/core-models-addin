/**
 * Excel utility functions for common operations
 */

/**
 * Convert zero-based (row, column) to A1 address, e.g. (0,0) -> "A1"
 * @param row Zero-based row index
 * @param column Zero-based column index
 * @returns Excel A1 notation address
 */
export function toA1(row: number, column: number): string {
  return `${columnToLetter(column)}${row + 1}`;
}

/**
 * Convert a 0-based column index to Excel column letters
 * @param column Zero-based column index (0 = A, 1 = B, etc.)
 * @returns Column letters (A, B, ..., Z, AA, AB, etc.)
 */
export function columnToLetter(column: number): string {
  let letters = '';
  let n = column;
  
  while (n >= 0) {
    letters = String.fromCharCode((n % 26) + 65) + letters;
    n = Math.floor(n / 26) - 1;
  }
  
  return letters;
}

/**
 * Convert Excel column letters to 0-based column index
 * @param columnLetters Column letters (A, B, ..., Z, AA, AB, etc.)
 * @returns Zero-based column index (A = 0, B = 1, etc.)
 */
export function letterToColumn(columnLetters: string): number {
  let column = 0;
  const letters = columnLetters.toUpperCase();
  
  for (let i = 0; i < letters.length; i++) {
    column = column * 26 + letters.charCodeAt(i) - 64; // 'A' is 65 in ASCII, 65-64=1
  }
  
  return column - 1; // Convert to 0-based
}

/**
 * Parse an Excel A1 reference into row and column indices
 * @param a1Reference A1-style reference (e.g., "A1", "BC123")
 * @returns Object with zero-based row and column indices
 */
export function parseA1(a1Reference: string): { row: number; column: number } {
  // Match letters followed by numbers
  const match = a1Reference.match(/^([A-Za-z]+)([0-9]+)$/);
  if (!match) {
    throw new Error(`Invalid A1 reference: ${a1Reference}`);
  }
  
  const columnLetters = match[1];
  const rowNumber = parseInt(match[2], 10);
  
  return {
    row: rowNumber - 1, // Convert to 0-based
    column: letterToColumn(columnLetters)
  };
}

/**
 * Parse a cell reference into sheet name and address
 * @param reference The cell reference (e.g. "Sheet1!A1")
 * @returns An object with sheet name and address
 */
export function parseReference(reference: string): { sheet: string; address: string } {
  const parts = reference.split('!');
  if (parts.length !== 2) {
    throw new Error(`Invalid cell reference: ${reference}`);
  }
  
  return {
    sheet: parts[0],
    address: parts[1]
  };
}

/**
 * Create a fully qualified reference from sheet name and address
 * @param sheet Sheet name
 * @param address Cell address in A1 notation
 * @returns Fully qualified reference (e.g., "Sheet1!A1")
 */
export function createReference(sheet: string, address: string): string {
  return `${sheet}!${address}`;
}

/**
 * Create a range address from top-left and bottom-right cells
 * @param topLeft Top-left cell in A1 notation
 * @param bottomRight Bottom-right cell in A1 notation
 * @returns Range address (e.g., "A1:B2")
 */
export function createRangeAddress(topLeft: string, bottomRight: string): string {
  return `${topLeft}:${bottomRight}`;
}
