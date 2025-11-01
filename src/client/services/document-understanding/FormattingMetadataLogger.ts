// File: src/client/services/document-understanding/FormattingMetadataLogger.ts
import * as path from 'path';
import * as fs from 'fs';
import { FormattingMetadataExtractor } from './FormattingMetadataExtractor';
import { WorkbookFormattingMetadata } from './FormattingModels';

/**
 * Utility class for logging formatting metadata to files
 */
export class FormattingMetadataLogger {
  /**
   * Extracts formatting metadata and saves it to a file
   * @param outputDir Directory to save the metadata file (defaults to data/extracted_metadata)
   * @returns The extracted metadata
   */
  public static async extractAndSaveMetadata(
    outputDir: string = path.join(process.cwd(), 'data', 'extracted_metadata')
  ): Promise<WorkbookFormattingMetadata> {
    try {
      // Create output directory if it doesn't exist
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
      }

      // Generate filename with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const filePath = path.join(outputDir, `metadata_${timestamp}.json`);

      // Extract metadata using the existing extractor
      const extractor = new FormattingMetadataExtractor();
      const metadata = await extractor.extractFormattingMetadata();

      // Add extraction timestamp
      const metadataWithTimestamp = {
        ...metadata,
        _extractedAt: new Date().toISOString()
      };

      // Save to file
      fs.writeFileSync(
        filePath,
        JSON.stringify(metadataWithTimestamp, null, 2),
        'utf-8'
      );

      console.log(`Metadata saved to: ${filePath}`);
      return metadata;
    } catch (error) {
      console.error('Error saving metadata to file:', error);
      throw error;
    }
  }
}