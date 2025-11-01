import * as fs from 'fs';
import * as path from 'path';

/**
 * Utility class for file operations
 */
export class FileUtils {
  /**
   * Ensures a directory exists, creating it if necessary
   * @param dirPath Path to the directory
   */
  private static ensureDirectoryExists(dirPath: string): void {
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
  }

  /**
   * Saves data to a JSON file
   * @param data Data to save (will be converted to JSON)
   * @param filePath Full path to the file including filename and extension
   */
  public static saveJsonToFile<T>(data: T, filePath: string): void {
    try {
      // Ensure directory exists
      const dirPath = path.dirname(filePath);
      this.ensureDirectoryExists(dirPath);
      
      // Write data to file with pretty-printing
      const jsonString = JSON.stringify(data, null, 2);
      fs.writeFileSync(filePath, jsonString, 'utf8');
      
      console.log(`Successfully saved data to ${filePath}`);
    } catch (error) {
      console.error(`Error saving JSON to file ${filePath}:`, error);
    }
  }

  /**
   * Generates a timestamp string for use in filenames
   * @returns Timestamp string in format YYYYMMDD_HHMMSS
   */
  public static getTimestampString(): string {
    const now = new Date();
    return [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, '0'),
      String(now.getDate()).padStart(2, '0'),
      '_',
      String(now.getHours()).padStart(2, '0'),
      String(now.getMinutes()).padStart(2, '0'),
      String(now.getSeconds()).padStart(2, '0')
    ].join('');
  }
}
