import { ClientWorkbookStateManager } from './ClientWorkbookStateManager';
import { MetadataChunk } from '../../models/CommandModels';

export class WorkbookMetadataLogger {
  /**
   * Logs the current workbook metadata to the console
   */
  public static async saveWorkbookMetadata(): Promise<void> {
    try {
      const stateManager = new ClientWorkbookStateManager();
      const workbookState = await stateManager.getCachedOrCaptureState();
      
      const metadataCache = (stateManager as any).metadataCache as {
        getAllChunks: () => MetadataChunk[];
        getWorkbookVersion: () => string;
      };

      const chunks = metadataCache.getAllChunks();
      
      const metadata = {
        extractedAt: new Date().toISOString(),
        workbookVersion: metadataCache.getWorkbookVersion?.(),
        chunksCount: chunks.length,
        chunks: chunks.map(chunk => ({
          id: chunk.id,
          type: chunk.type,
          etag: chunk.etag,
          lastCaptured: chunk.lastCaptured
        })),
        workbookState: {
          sheets: workbookState.sheets.map(sheet => sheet.name),
          activeSheet: workbookState.activeSheet
        }
      };

      console.log('WORKBOOK METADATA:', JSON.stringify(metadata, null, 2));
    } catch (error) {
      console.error('Error logging workbook metadata:', error);
      throw error;
    }
  }
}