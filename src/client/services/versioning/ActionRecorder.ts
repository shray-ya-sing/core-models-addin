/**
 * Action Recorder
 * 
 * Provides functionality to record Excel operations before they are executed,
 * capturing the before state to enable precise undo operations.
 */

import { v4 as uuidv4 } from 'uuid';
import { ExcelOperation, ExcelOperationType } from '../../models/ExcelOperationModels';
import { AffectedRange, BeforeState } from '../../models/VersionModels';
import { VersionHistoryService } from './VersionHistoryService';

/**
 * Service for recording Excel operations and their before states
 */
export class ActionRecorder {
  private versionHistoryService: VersionHistoryService;
  private recentlyRecordedOperations: Map<string, number> = new Map();
  private operationExpiryTime: number = 5000; // 5 seconds
  private recordedOperationIds: Set<string> = new Set();
  private recordedOperationFingerprints: Set<string> = new Set();
  private readonly OPERATION_TRACKING_WINDOW_MS = 5000; // 5 seconds window to prevent duplicates
  
  /**
   * Constructor that accepts a VersionHistoryService instance
   * @param versionHistoryService The version history service to use
   */
  constructor(versionHistoryService?: VersionHistoryService) {
    // Use the provided instance or get the singleton instance
    this.versionHistoryService = versionHistoryService || VersionHistoryService.getInstance();
    
    // Ensure the service is initialized
    if (!versionHistoryService) {
      this.versionHistoryService.initialize();
    }
  }
  
  /**
   * Clean up old operation IDs from the tracking map
   * @private
   */
  private cleanupOldOperations(): void {
    const now = Date.now();
    const expiredIds: string[] = [];
    
    // Identify expired operation IDs
    this.recentlyRecordedOperations.forEach((timestamp, id) => {
      if (now - timestamp > this.OPERATION_TRACKING_WINDOW_MS) {
        expiredIds.push(id);
      }
    });
    
    // Remove expired IDs
    expiredIds.forEach(id => {
      this.recentlyRecordedOperations.delete(id);
    });
    
    if (expiredIds.length > 0) {
      console.log(`🧹 [ActionRecorder] Cleaned up ${expiredIds.length} expired operation IDs`);
    }
  }
  
  /**
   * Creates a unique fingerprint for an operation based on its properties
   * This helps detect duplicate operations even if they have different IDs
   * @param operation The Excel operation to create a fingerprint for
   * @returns A string fingerprint that uniquely identifies this operation
   * @private
   */
  private createOperationFingerprint(operation: ExcelOperation): string {
    // Extract key properties that identify an operation
    const opType = operation.op || 'unknown';
    
    // Safely extract properties based on operation type
    let targetOrRange = '';
    let valueOrFormula = '';
    let formatInfo = '';
    
    // Get properties safely using type assertion
    const op = operation as any;
    
    // Extract common properties based on operation type string
    if (opType === 'set_value') {
      targetOrRange = op.target || '';
      valueOrFormula = op.value !== undefined ? String(op.value) : '';
    } 
    else if (opType === 'add_formula') {
      targetOrRange = op.target || '';
      valueOrFormula = op.formula || '';
    }
    else if (opType === 'format_range') {
      targetOrRange = op.range || '';
      formatInfo = op.style || '';
    }
    else if (opType === 'create_sheet') {
      valueOrFormula = op.name || '';
    }
    else if (['delete_sheet', 'rename_sheet', 'set_active_sheet'].includes(opType)) {
      targetOrRange = op.name || op.sheet || '';
    }
    else {
      // For other operations, try to get common properties
      targetOrRange = op.target || op.range || op.sheet || '';
      valueOrFormula = op.value || op.formula || op.name || '';
      
      // Try to extract format information if available
      if (op.format) {
        try {
          formatInfo = typeof op.format === 'object' ? JSON.stringify(op.format) : String(op.format);
        } catch (e) {
          // If JSON stringification fails, use a simpler approach
          formatInfo = 'has-format';
        }
      }
    }
    
    // Create a fingerprint string that combines these properties
    // We don't include the timestamp in the fingerprint to ensure identical operations are detected
    // even if they're a few milliseconds apart
    const fingerprint = `${opType}:${targetOrRange}:${valueOrFormula}:${formatInfo}`;
    
    return fingerprint;
  }
  
  /**
   * Record an Excel operation before it is executed
   * @param context The Excel context
   * @param workbookId The workbook ID
   * @param operation The operation to record
   * @returns A promise that resolves when the operation is recorded
   */
  public async recordOperation(
    context: Excel.RequestContext,
    workbookId: string,
    operation: ExcelOperation
  ): Promise<string> {
    const startTime = performance.now();
    const opType = operation.op || 'unknown';
    
    try {
      // Clean up old operation IDs
      this.cleanupOldOperations();
      
      // Generate an ID for the operation if it doesn't have one
      if (!operation.id) {
        operation.id = uuidv4();
        console.log(`🆔 [ActionRecorder] Generated ID for operation: ${operation.id}`);
      }
      
      // Create a fingerprint for this operation to detect duplicates even with different IDs
      const fingerprint = this.createOperationFingerprint(operation);
      
      // Check if this operation has already been recorded by ID or fingerprint
      if (this.recordedOperationIds.has(operation.id)) {
        console.log(`♻️ [ActionRecorder] Skipping already recorded operation by ID: ${operation.id}`);
        return operation.id;
      }
      
      if (this.recordedOperationFingerprints.has(fingerprint)) {
        console.log(`♻️ [ActionRecorder] Skipping already recorded operation by fingerprint: ${fingerprint}`);
        return operation.id;
      }
      
      if (this.recentlyRecordedOperations.has(operation.id) || this.recentlyRecordedOperations.has(fingerprint)) {
        console.log(`⚠️ [ActionRecorder] Skipping duplicate operation with ID: ${operation.id} (fingerprint: ${fingerprint})`);
        // Return the existing operation ID without re-recording
        return operation.id;
      }
      
      console.log(`🔄 [ActionRecorder] Starting capture for operation type: ${opType} on workbook: ${workbookId}`);
      console.log(`📝 [ActionRecorder] Operation details:`, {
        id: operation.id,
        type: opType,
        params: Object.keys(operation).filter(k => k !== 'op' && k !== 'id').map(k => `${k}: ${typeof operation[k] === 'object' ? '[object]' : operation[k]}`)
      });
      
      // Identify affected ranges
      const affectedRanges = this.identifyAffectedRanges(operation);
      console.log(`📊 [ActionRecorder] Identified ${affectedRanges.length} affected ranges:`, 
        affectedRanges.map(r => `${r.sheetName}!${r.range || '[sheet-level]'}`).join(', '));
      
      // Capture the before state
      console.log(`📷 [ActionRecorder] Capturing before state...`);
      const beforeState = await this.captureBeforeState(context, affectedRanges);
      
      // Log capture statistics with more details
      const hasValues = beforeState.values && 
        Array.isArray(beforeState.values) && 
        beforeState.values.length > 0;
      
      const hasFormulas = beforeState.formulas && 
        Array.isArray(beforeState.formulas) && 
        beforeState.formulas.length > 0;
      
      const formatCount = beforeState.formats ? beforeState.formats.length : 0;
      const sheetPropertiesCount = Object.keys(beforeState.sheetProperties || {}).length;
      
      // Sample of captured data for debugging
      const valueSample = hasValues && beforeState.values.length > 0 ? 
        JSON.stringify(beforeState.values[0]).substring(0, 100) + '...' : 'none';
      const formulaSample = hasFormulas && beforeState.formulas.length > 0 ? 
        JSON.stringify(beforeState.formulas[0]).substring(0, 100) + '...' : 'none';
      
      const endTime = performance.now();
      console.log(`✅ [ActionRecorder] Capture successful for ${opType}:`, {
        duration: `${(endTime - startTime).toFixed(2)}ms`,
        stats: {
          hasValues,
          valueRows: hasValues ? beforeState.values.length : 0,
          valueSample,
          hasFormulas,
          formulaRows: hasFormulas ? beforeState.formulas.length : 0,
          formulaSample,
          formatCount,
          sheetPropertiesCount
        },
        affectedRanges: affectedRanges.map(r => `${r.sheetName}!${r.range || '[sheet-level]'}`)
      });
      
      // Record the action
      console.log(`📤 [ActionRecorder] Sending to VersionHistoryService for recording...`);
      const actionId = this.versionHistoryService.recordAction(
        workbookId,
        operation,
        beforeState,
        affectedRanges
      );
      
      // Only after successful recording, add both the operation ID and fingerprint to the tracking map
      // This prevents duplicate recordings of the same operation
      this.recentlyRecordedOperations.set(operation.id, Date.now());
      this.recentlyRecordedOperations.set(fingerprint, Date.now());
      
      // Also add to our permanent tracking sets
      this.recordedOperationIds.add(operation.id);
      this.recordedOperationFingerprints.add(fingerprint);
      
      console.log(`💾 [ActionRecorder] Action successfully recorded with ID: ${actionId}`);
      return actionId;
    } catch (error) {
      const endTime = performance.now();
      console.error(`❌ [ActionRecorder] Error capturing operation ${opType}:`, {
        error: error.message,
        stack: error.stack,
        duration: `${(endTime - startTime).toFixed(2)}ms`,
        operation: JSON.stringify(operation).substring(0, 200), // Truncate for readability
        workbookId
      });
      
      // Return a dummy ID - we don't want to fail the main operation if recording fails
      return uuidv4();
    }
  }
  
  /**
   * Identify ranges affected by an operation
   * @param operation The Excel operation
   * @returns Array of affected ranges
   */
  private identifyAffectedRanges(operation: ExcelOperation): AffectedRange[] {
    const affectedRanges: AffectedRange[] = [];
    
    // Extract sheet name and range based on operation type
    const opType = operation.op as string;
    
    switch (opType) {
      case ExcelOperationType.SET_VALUE:
        // For setValue operations, the target is the cell reference
        if ('target' in operation) {
          const [sheetName, cellRef] = this.parseReference(operation.target as string);
          affectedRanges.push({
            sheetName,
            range: cellRef,
            type: 'cell'
          });
        }
        break;
        
      case ExcelOperationType.ADD_FORMULA:
        // For formula operations, the target is the cell reference
        if ('target' in operation) {
          const [formulaSheetName, formulaCellRef] = this.parseReference(operation.target as string);
          affectedRanges.push({
            sheetName: formulaSheetName,
            range: formulaCellRef,
            type: 'cell'
          });
        }
        break;
        
      case ExcelOperationType.FORMAT_RANGE:
      case ExcelOperationType.CLEAR_RANGE:
      case ExcelOperationType.SORT_RANGE:
      case ExcelOperationType.FILTER_RANGE:
      case ExcelOperationType.MERGE_CELLS:
      case ExcelOperationType.UNMERGE_CELLS:
        // For range operations, the range property contains the reference
        if ('range' in operation) {
          const [rangeSheetName, rangeRef] = this.parseReference(operation.range);
          affectedRanges.push({
            sheetName: rangeSheetName,
            range: rangeRef,
            type: 'range'
          });
        }
        break;
        
      case ExcelOperationType.CREATE_SHEET:
      case ExcelOperationType.DELETE_SHEET:
      case ExcelOperationType.RENAME_SHEET:
        // For sheet operations, the name property contains the sheet name
        if ('name' in operation) {
          affectedRanges.push({
            sheetName: operation.name,
            range: '',
            type: 'sheet'
          });
        }
        break;
        
      case ExcelOperationType.CREATE_TABLE:
        // For table operations, the range property contains the reference
        if ('range' in operation) {
          const [tableSheetName, tableRange] = this.parseReference(operation.range);
          affectedRanges.push({
            sheetName: tableSheetName,
            range: tableRange,
            type: 'table'
          });
        }
        break;
        
      case ExcelOperationType.CREATE_CHART:
        // For chart operations, the range property contains the data range
        if ('range' in operation) {
          const [chartSheetName, chartRange] = this.parseReference(operation.range);
          affectedRanges.push({
            sheetName: chartSheetName,
            range: chartRange,
            type: 'chart'
          });
        }
        break;
        
      case ExcelOperationType.COPY_RANGE:
        // For copy operations, both source and destination are affected
        if ('source' in operation && 'destination' in operation) {
          const [sourceSheetName, sourceRange] = this.parseReference(operation.source);
          const [destSheetName, destRange] = this.parseReference(operation.destination);
          
          affectedRanges.push({
            sheetName: sourceSheetName,
            range: sourceRange,
            type: 'range'
          });
          
          affectedRanges.push({
            sheetName: destSheetName,
            range: destRange,
            type: 'range'
          });
        }
        break;
        
      case ExcelOperationType.COMPOSITE_OPERATION:
      case ExcelOperationType.BATCH_OPERATION:
        // For composite operations, we need to analyze each sub-operation
        if ('operations' in operation) {
          for (const subOp of operation.operations) {
            const subRanges = this.identifyAffectedRanges(subOp);
            affectedRanges.push(...subRanges);
          }
        } else if ('subOperations' in operation) {
          for (const subOp of operation.subOperations) {
            const subRanges = this.identifyAffectedRanges(subOp);
            affectedRanges.push(...subRanges);
          }
        }
        break;
    }
    
    return affectedRanges;
  }
  
  /**
   * Capture the state of affected ranges before an operation is executed
   * @param context The Excel context
   * @param affectedRanges The ranges affected by the operation
   * @returns The before state
   */
  private async captureBeforeState(
    context: Excel.RequestContext,
    affectedRanges: AffectedRange[]
  ): Promise<BeforeState> {
    const captureStartTime = performance.now();
    const beforeState: BeforeState = {
      values: [],
      formulas: [],
      formats: [],
      sheetProperties: {}
    };
    
    try {
      // Group ranges by sheet
      const rangesBySheet = new Map<string, AffectedRange[]>();
      
      for (const range of affectedRanges) {
        if (!rangesBySheet.has(range.sheetName)) {
          rangesBySheet.set(range.sheetName, []);
        }
        rangesBySheet.get(range.sheetName)?.push(range);
      }
      
      console.log(`🗂️ Processing ${rangesBySheet.size} sheets for capture`);
      
      // Process each sheet
      for (const [sheetName, ranges] of rangesBySheet.entries()) {
        console.log(`📋 Processing sheet "${sheetName}" with ${ranges.length} ranges`);
        
        // Get the worksheet
        let worksheet: Excel.Worksheet;
        try {
          worksheet = context.workbook.worksheets.getItem(sheetName);
          console.log(`✓ Found sheet "${sheetName}"`);
        } catch (error) {
          console.warn(`⚠️ Sheet "${sheetName}" not found, skipping capture: ${error.message}`);
          continue;
        }
        
        // Capture sheet properties if any sheet-level operations
        if (ranges.some(r => r.type === 'sheet')) {
          console.log(`📝 Loading sheet properties for "${sheetName}"`);
          worksheet.load(['name', 'position', 'visibility']);
          beforeState.sheetProperties[sheetName] = {};
        }
        
        // Capture cell and range values
        let loadedRanges = 0;
        let skippedRanges = 0;
        let errorRanges = 0;
        
        for (const rangeInfo of ranges) {
          if (rangeInfo.type === 'sheet') {
            console.log(`⏩ Skipping range capture for sheet-level operation on "${sheetName}"`);
            skippedRanges++;
            continue; // Skip sheet-level operations for range capture
          }
          
          if (!rangeInfo.range) {
            console.log(`⏩ Skipping range capture for empty range on "${sheetName}"`);
            skippedRanges++;
            continue; // Skip if no range specified
          }
          
          try {
            console.log(`🔍 Loading range "${rangeInfo.range}" in sheet "${sheetName}"`);
            const excelRange = worksheet.getRange(rangeInfo.range);
            excelRange.load(['values', 'formulas', 'format']);
            loadedRanges++;
            
            // We'll populate these after context.sync()
            beforeState.values = [];
            beforeState.formulas = [];
            beforeState.formats = [];
          } catch (error) {
            console.warn(`⚠️ Error loading range "${rangeInfo.range}" in sheet "${sheetName}": ${error.message}`);
            errorRanges++;
          }
        }
        
        console.log(`📊 Range loading summary for "${sheetName}": ${loadedRanges} loaded, ${skippedRanges} skipped, ${errorRanges} errors`);
        
        // Sync to load the properties
        console.log(`🔄 Syncing context to load properties for sheet "${sheetName}"...`);
        const syncStartTime = performance.now();
        try {
          await context.sync();
          const syncEndTime = performance.now();
          console.log(`✓ Context sync completed in ${(syncEndTime - syncStartTime).toFixed(2)}ms`);
        } catch (syncError) {
          console.error(`❌ Error syncing context for sheet "${sheetName}": ${syncError.message}`);
          continue; // Skip this sheet if sync fails
        }
        
        // Now extract the loaded properties
        if (beforeState.sheetProperties[sheetName]) {
          try {
            beforeState.sheetProperties[sheetName] = {
              name: worksheet.name,
              position: worksheet.position,
              visibility: worksheet.visibility
            };
            console.log(`✓ Captured sheet properties for "${sheetName}"`);
          } catch (propError) {
            console.warn(`⚠️ Error extracting sheet properties for "${sheetName}": ${propError.message}`);
          }
        }
        
        // Extract range values after sync
        let capturedRanges = 0;
        let captureErrors = 0;
        
        for (const rangeInfo of ranges) {
          if (rangeInfo.type === 'sheet' || !rangeInfo.range) {
            continue;
          }
          
          try {
            console.log(`📥 Extracting data from range "${rangeInfo.range}" in sheet "${sheetName}"`);
            const excelRange = worksheet.getRange(rangeInfo.range);
            
            // Load all the properties we need before accessing them
            excelRange.load([
              'values', 
              'formulas',
              'format/font/name',
              'format/font/size',
              'format/font/bold',
              'format/font/italic',
              'format/font/underline',
              'format/font/color',
              'format/fill/color',
              'format/horizontalAlignment',
              'format/verticalAlignment'
            ]);
            
            // Sync to get the loaded properties
            await context.sync();
            console.log(`✓ Successfully loaded range properties for "${rangeInfo.range}" in sheet "${sheetName}"`);
            
            // Now we can safely access the properties
            beforeState.values = excelRange.values;
            beforeState.formulas = excelRange.formulas;
            
            // Store formatting (simplified)
            const format = {
              fontFamily: excelRange.format.font.name,
              fontSize: excelRange.format.font.size,
              bold: excelRange.format.font.bold,
              italic: excelRange.format.font.italic,
              underline: excelRange.format.font.underline,
              fontColor: excelRange.format.font.color,
              fillColor: excelRange.format.fill.color,
              horizontalAlignment: excelRange.format.horizontalAlignment,
              verticalAlignment: excelRange.format.verticalAlignment
            };
            
            beforeState.formats.push(format);
            capturedRanges++;
            
            // Log a sample of the captured data for debugging
            if (beforeState.values && Array.isArray(beforeState.values) && beforeState.values.length > 0) {
              const sampleRow = beforeState.values[0];
              console.log(`📊 Captured data for "${rangeInfo.range}" in sheet "${sheetName}": ${JSON.stringify(sampleRow).substring(0, 100)}${sampleRow.length > 100 ? '...' : ''}`);
            }
          } catch (error) {
            console.warn(`⚠️ Error capturing range "${rangeInfo.range}" in sheet "${sheetName}": ${error.message}`);
            captureErrors++;
          }
        }
        
        console.log(`📊 Range capture summary for "${sheetName}": ${capturedRanges} captured, ${captureErrors} errors`);
      }
      
      const captureEndTime = performance.now();
      console.log(`✅ Before state capture completed in ${(captureEndTime - captureStartTime).toFixed(2)}ms`);
      
      return beforeState;
    } catch (error) {
      const captureEndTime = performance.now();
      console.error(`❌ Error capturing before state: ${error.message}`, {
        duration: `${(captureEndTime - captureStartTime).toFixed(2)}ms`,
        stack: error.stack
      });
      return beforeState; // Return empty state on error
    }
  }
  
  /**
   * Parse a cell reference into sheet name and cell reference
   * @param reference The cell reference (e.g., "Sheet1!A1" or "A1")
   * @returns Tuple of [sheetName, cellReference]
   */
  private parseReference(reference: string): [string, string] {
    // Check if the reference includes a sheet name
    if (reference.includes('!')) {
      const [sheetName, cellRef] = reference.split('!');
      return [sheetName, cellRef];
    }
    
    // If no sheet name, assume it's on the active sheet
    return ['', reference];
  }
}
