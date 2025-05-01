// PART 1: IMPORTS AND CLASS DEFINITION
import { MetadataChunk, SheetState, WorkbookState } from '../models/CommandModels';
import { RangeDetectionResult } from '../models/RangeModels';
import { RangeDetector } from './RangeDetector';
import { SpreadsheetChunkCompressor } from './SpreadsheetChunkCompressor';
import { WorkbookMetadataCache } from './WorkbookMetadataCache';
import { RangeDependencyAnalyzer } from './RangeDependencyAnalyzer';

// Office.js import - this should be available globally in the Excel add-in environment
declare const Excel: any;

/**
 * Client-side workbook state manager for Excel
 * Handles sheet-level and range-level granularity for workbook state capture
 * with integrated range detection capabilities
 */
export class ClientWorkbookStateManager {
  // Cache for workbook state to avoid redundant captures
  private cachedState: WorkbookState | null = null;
  private lastCaptureTime: number = 0; // timestamp in milliseconds
  private cacheTimeoutMs: number = 5000; // 5 seconds cache timeout by default
  private cacheEnabled: boolean = true;
  
  // Chunk-based metadata components
  private metadataCache: WorkbookMetadataCache;
  private chunkCompressor: SpreadsheetChunkCompressor;
  private dependencyAnalyzer: RangeDependencyAnalyzer;
  private rangeDetector: RangeDetector;
  
  private enableRangeDetection: boolean = true; // Set to false to disable range-level granularity
  private activeSheetName: string = '';
  
  constructor(cacheTimeoutMs?: number) {
    if (cacheTimeoutMs !== undefined) {
      this.cacheTimeoutMs = cacheTimeoutMs;
    }
    
    // Initialize chunk-based components
    this.metadataCache = new WorkbookMetadataCache();
    this.chunkCompressor = new SpreadsheetChunkCompressor();
    this.dependencyAnalyzer = this.metadataCache.getDependencyAnalyzer();
    this.rangeDetector = new RangeDetector();
  }
  
  /**
   * Set up event listeners to detect workbook changes
   */
  public async setupChangeListeners(): Promise<void> {
    try {
      // Use Excel API to set up event handlers for worksheet changes
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        
        // Listen for sheet-level change events
        const worksheets = workbook.worksheets;
        worksheets.load('items/name');
        worksheets.onChanged.add(this.handleWorkbookChange.bind(this));
        worksheets.onAdded.add(this.handleWorkbookChange.bind(this));
        worksheets.onDeleted.add(this.handleWorkbookChange.bind(this));
        worksheets.onActivated.add(this.handleWorkbookChange.bind(this));
        
        await context.sync();
        console.log('%c Workbook state cache: Event listeners registered', 'color: #2980b9');
      });
    } catch (error) {
      console.error('Error setting up workbook listeners:', error);
      // If we can't set up listeners, disable caching to ensure fresh data
      this.cacheEnabled = false;
    }
  }
  
  /**
   * Handle workbook change events to invalidate cache
   */
  public handleWorkbookChange(): void {
    console.log('%c Workbook changed, invalidating cache', 'color: #e74c3c');
    this.invalidateCache();
  }
  
  /**
   * Set the active sheet name
   * @param name Active sheet name
   */
  public setActiveSheetName(name: string): void {
    this.activeSheetName = name;
  }
  
  /**
   * Get the active sheet name
   * @returns Active sheet name
   */
  public getActiveSheetName(): string {
    return this.activeSheetName;
  }
  
  // PART 2: WORKBOOK STATE MANAGEMENT METHODS
  /**
   * Get cached workbook state or capture a new one if needed
   * @param forceRefresh Force a refresh of the workbook state
   * @returns The workbook state (either cached or freshly captured)
   */
  public async getCachedOrCaptureState(forceRefresh = false): Promise<WorkbookState> {
    const currentTime = Date.now();
    const cacheAge = currentTime - this.lastCaptureTime;
    
    // Use cache if it exists, is not too old, caching is enabled, and refresh is not forced
    if (this.cachedState && cacheAge < this.cacheTimeoutMs && this.cacheEnabled && !forceRefresh) {
      console.log(`%c Using cached workbook state (age: ${cacheAge}ms)`, 'color: #27ae60');
      return this.cachedState;
    }
    
    // Otherwise, refresh state
    console.log('%c Capturing fresh workbook state', 'color: #3498db');
    this.cachedState = await this.captureWorkbookState();
    this.lastCaptureTime = currentTime;
    return this.cachedState;
  }
  
  /**
   * Invalidate the workbook state cache
   * @param operationType Optional operation type or comma-separated list of operation types to selectively invalidate based on operation
   */
  public invalidateCache(operationType?: string): void {
    // If no operation type is provided, invalidate the full cache
    if (!operationType) {
      console.log('%c Invalidating workbook state cache (no operation type specified)', 'color: #e74c3c');
      this.cachedState = null;
      this.lastCaptureTime = 0;
      this.metadataCache.invalidateAllChunks();
      return;
    }
    
    // Handle multiple operation types (comma-separated)
    const operationTypes = operationType.includes(',') 
      ? operationType.split(',').map(op => op.trim()) 
      : [operationType];
    
    // Check if any of the operation types require cache invalidation
    const requiresInvalidation = operationTypes.some(op => this.requiresFullCacheInvalidation(op));
    
    if (requiresInvalidation) {
      console.log('%c Invalidating workbook state cache for operations: ' + operationTypes.join(', '), 'color: #e74c3c');
      this.cachedState = null;
      this.lastCaptureTime = 0;
      
      // Also invalidate chunks in the metadata cache
      this.metadataCache.invalidateAllChunks();
    } else {
      console.log(`%c Operations [${operationTypes.join(', ')}] do not require cache invalidation`, 'color: #2ecc71');
    }
  }
  
  /**
   * Determines if an operation type requires full cache invalidation
   * @param operationType The operation type to check
   * @returns True if the operation modifies data and requires cache invalidation
   */
  private requiresFullCacheInvalidation(operationType: string): boolean {
    // Normalize operation type to lowercase for case-insensitive comparison
    const normalizedOp = operationType.toLowerCase();
    
    // Operations that only modify UI/display settings and don't need cache invalidation
    const uiOnlyOperations = [
      'set_gridlines',
      'set_headers',
      'set_zoom',
      'set_freeze_panes',
      'set_visible',
      'set_active_sheet',
      'set_print_area',
      'set_print_orientation',
      'set_print_margins',
      'format_chart',
      'format_chart_axis',
      'format_chart_series'
    ];
    
    // Check if the operation is in the UI-only list
    if (uiOnlyOperations.includes(normalizedOp)) {
      console.log(`Operation ${operationType} is UI-only and doesn't require cache invalidation`);
      return false;
    }
    
    // Operations that modify data and require cache invalidation
    const dataModifyingOperations = [
      'set_value',
      'add_formula',
      'create_chart',
      'format_range',
      'clear_range',
      'create_table',
      'sort_range',
      'filter_range',
      'create_sheet',
      'delete_sheet',
      'rename_sheet',
      'copy_range',
      'merge_cells',
      'unmerge_cells',
      'conditional_format',
      'add_comment'
    ];
    
    // Check if the operation is in the data-modifying list
    const requiresInvalidation = dataModifyingOperations.includes(normalizedOp);
    
    // If not explicitly listed in either category, default to requiring invalidation (safer)
    if (!requiresInvalidation && !uiOnlyOperations.includes(normalizedOp)) {
      console.log(`Operation ${operationType} not recognized, defaulting to cache invalidation for safety`);
      return true;
    }
    
    return requiresInvalidation;
  }
  
  /**
   * Set whether caching is enabled
   * @param enabled Whether caching is enabled
   */
  public setCacheEnabled(enabled: boolean): void {
    this.cacheEnabled = enabled;
    console.log(`%c Workbook state caching ${enabled ? 'enabled' : 'disabled'}`, 'color: #3498db');
    
    // If disabling, invalidate the current cache
    if (!enabled) {
      this.invalidateCache();
    }
  }
  
  /**
   * Set whether range detection is enabled
   * @param enabled Whether range detection is enabled
   */
  public setRangeDetectionEnabled(enabled: boolean): void {
    this.enableRangeDetection = enabled;
    console.log(`%c Range detection ${enabled ? 'enabled' : 'disabled'}`, 'color: #3498db');
  }
  
  /**
   * Get the chunk compressor
   * @returns The chunk compressor instance
   */
  public getChunkCompressor(): SpreadsheetChunkCompressor {
    return this.chunkCompressor;
  }
  
  /**
   * Get the metadata cache
   * @returns The metadata cache instance
   */
  public getMetadataCache(): WorkbookMetadataCache {
    return this.metadataCache;
  }
  
  /**
   * Get the dependency analyzer
   * @returns The dependency analyzer instance
   */
  public getDependencyAnalyzer(): RangeDependencyAnalyzer {
    return this.dependencyAnalyzer;
  }
  
  /**
   * Get the range detector
   * @returns The range detector instance
   */
  public getRangeDetector(): RangeDetector {
    return this.rangeDetector;
  }

  // PART 3: WORKBOOK AND SHEET STATE CAPTURE
  /**
   * Capture the current state of the Excel workbook
   * @returns The workbook state
   */
  public async captureWorkbookState(): Promise<WorkbookState> {
    try {
      return await Excel.run(async (context) => {
        // Get all worksheets
        
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        
        // Get active worksheet
        const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
        activeWorksheet.load('name');
        
        await context.sync();
        
        // Set active sheet name for future reference
        this.setActiveSheetName(activeWorksheet.name);
        
        // Create promises for capturing each sheet's state
        // For Phase 1, we use parallel capture for better performance
        const sheetPromises = worksheets.items.map(worksheet => {
          return this.captureSheetState(worksheet.name);
        });
        
        // Wait for all sheet states to be captured
        const sheets = await Promise.all(sheetPromises);
        
        // Add each sheet to the metadata cache with dependency analysis
        this.processSheetStatesForCache(sheets);
        
        // Return workbook state with all sheet states
        return {
          sheets,
          activeSheet: activeWorksheet.name
        };
      });
    } catch (error) {
      console.error('Error capturing workbook state:', error);
      throw error;
    }
  }
  
  /**
   * Process captured sheet states to add to metadata cache with dependency analysis
   * @param sheets Array of captured sheet states
   */
  private processSheetStatesForCache(sheets: SheetState[]): void {
    console.log(`%c Processing ${sheets.length} sheets for metadata cache`, 'color: #3498db');
    
    // Track successful and failed sheets
    let successCount = 0;
    let failCount = 0;
    
    // Process each sheet and add to cache
    for (const sheet of sheets) {
      try {
        // Compress the sheet into a chunk
        const chunk = this.chunkCompressor.compressSheetToChunk(sheet);
        
        // Add the chunk to the cache with dependency analysis
        this.metadataCache.addChunkWithDependencyAnalysis(chunk);
        
        // If range detection is enabled, also capture ranges for this sheet
        if (this.enableRangeDetection) {
          this.captureRanges(sheet);
        }
        
        console.log(`%c Cached and analyzed sheet: ${sheet.name}`, 'color: #2ecc71');
        successCount++;
      } catch (error) {
        console.error(`%c Error processing sheet "${sheet.name}": ${error.message}`, 'color: #e74c3c');
        failCount++;
        
        // Create a simplified chunk with just the basic info to avoid completely missing this sheet
        try {
          // Create a basic chunk with minimal info
          const basicChunk = {
            id: `Sheet:${sheet.name}`,
            type: 'sheet' as 'sheet',
            etag: new Date().getTime().toString(), // Simple timestamp as etag
            payload: {
              name: sheet.name,
              summary: `Sheet ${sheet.name} (basic info only due to processing error)`,
              anchors: [],
              values: []
            },
            refs: [],
            lastCaptured: new Date()
          };
          
          this.metadataCache.setChunk(basicChunk);
          console.log(`%c Created simplified chunk for sheet: ${sheet.name}`, 'color: #f39c12');
        } catch (fallbackError) {
          console.error(`%c Failed to create even a basic chunk for ${sheet.name}:`, 'color: #e74c3c', fallbackError);
        }
      }
    }
    
    console.log(`%c Sheet processing complete: ${successCount} succeeded, ${failCount} had errors`, 
                successCount > 0 ? 'color: #2ecc71' : 'color: #e74c3c');
                
    // If all sheets failed, log a warning
    if (successCount === 0 && failCount > 0) {
      console.warn(`Failed to process any sheets successfully (${failCount} sheets had errors)`);
    }
  }

  /**
   * Capture sheet state for a specific worksheet
   * @param sheetName Name of the worksheet
   * @returns Promise with the sheet state
   */
  private async captureSheetState(sheetName: string): Promise<SheetState> {
    console.log(`Loading data for sheet: ${sheetName}`);
    
    try {
      return await Excel.run(async (context) => {
        // Get the worksheet from this context
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.load(['name']);
        
        // Load used range for the worksheet
        const usedRange = sheet.getUsedRange();
        usedRange.load(['address', 'columnCount', 'rowCount', 'formulas', 'values']);
        
        // Load tables if any
        const tables = sheet.tables;
        tables.load(['items/name']);
        
        // Load charts if any
        const charts = sheet.charts;
        charts.load('items');
        
        await context.sync();
        
        // If charts exist, load their detailed properties (name, position, size)
        if (charts.items.length > 0) {
          charts.items.forEach(chart => chart.load(['name', 'left', 'top', 'height', 'width']));
          await context.sync();
        }
        
        // Create sheet state object
        const sheetState: SheetState = {
          name: sheet.name,
          usedRange: {
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount
          },
          values: usedRange.values,
          formulas: usedRange.formulas
        };
        
        // Add tables if any
        if (tables.items.length > 0) {
          sheetState.tables = [];
          
          for (const table of tables.items) {
            try {
              // Load table details
              const tableObj = context.workbook.tables.getItem(table.name);
              tableObj.load(['headerRowRange']);
              const tableRange = tableObj.getRange();
              tableRange.load(['address']);
              await context.sync();
              
              // Get headers
              let headers: string[] = [];
              try {
                if (tableObj.headerRowRange) {
                  const headerRange = tableObj.getHeaderRowRange();
                  headerRange.load(['values']);
                  await context.sync();
                  headers = headerRange.values[0] as string[];
                }
              } catch (error) {
                console.warn(`Could not get headers for table ${table.name}:`, error);
              }
              
              sheetState.tables.push({
                name: table.name,
                range: tableRange.address,
                headers: headers
              });
            } catch (error) {
              console.warn(`Error processing table ${table.name}:`, error);
            }
          }
          
          console.log(`Added ${sheetState.tables.length} tables to sheet state for ${sheetName}`);
        }

        // Add charts if any
        if (charts.items.length > 0) {
          sheetState.charts = [];
          
          for (const chart of charts.items) {
            sheetState.charts.push({
              name: chart.name,
              type: 'chart', // Generic chart type
              range: `${chart.left},${chart.top},${chart.width},${chart.height}`
            });
          }
          
          console.log(`Added ${sheetState.charts.length} charts to sheet state for ${sheetName}`);
        }
        
        // Try to get named ranges
        try {
          const namedItems = context.workbook.names;
          // First load the collection itself
          namedItems.load('items');
          await context.sync();

          // Explicitly load name and value for each NamedItem
          namedItems.items.forEach(item => item.load(['name', 'value']));
          await context.sync();
          
          // Filter for named ranges in this sheet, guarding against unloaded or undefined values
          const namedRanges = namedItems.items
            .filter(item => {
              try {
                return typeof item.value === 'string' && item.value.includes(`${sheetName}!`);
              } catch {
                // Ignore items where value is not loaded or not a string
                return false;
              }
            })
            .map(item => ({
              name: item.name,
              value: item.value as string
            }));
          
          if (namedRanges.length > 0) {
            // Add named ranges to the sheet state
            (sheetState as any).namedRanges = namedRanges;
            console.log(`Added ${namedRanges.length} named ranges to sheet state for ${sheetName}`);
          }
        } catch (error) {
          console.warn(`Error loading named ranges for sheet ${sheetName}:`, error);
        }
        
        return sheetState;
      });
    } catch (error) {
      console.error(`Error capturing sheet state for ${sheetName}:`, error);
      // Return a minimal sheet state in case of error
      return {
        name: sheetName,
        usedRange: {
          rowCount: 0,
          columnCount: 0
        },
        values: [],
        formulas: []
      };
    }
  }
  
  /**
   * Capture ranges for a sheet and add them to the metadata cache
   * @param sheet The sheet state
   * @returns Array of range chunks that were captured
   */
  public captureRanges(sheet: SheetState): MetadataChunk[] {
    if (!this.enableRangeDetection) {
      return [];
    }
    
    try {
      console.log(`%c Detecting ranges in sheet: ${sheet.name}`, 'color: #9b59b6');
      
      // Detect ranges in the sheet
      const detectionResult: RangeDetectionResult = this.rangeDetector.detectRanges(sheet);
      
      if (detectionResult.ranges.length === 0) {
        console.log(`%c No significant ranges detected in sheet: ${sheet.name}`, 'color: #95a5a6');
        return [];
      }
      
      console.log(`%c Detected ${detectionResult.ranges.length} ranges in sheet: ${sheet.name}`, 'color: #2ecc71');
      
      // Create chunks for the detected ranges
      const rangeChunks = this.rangeDetector.createRangeChunks(sheet, detectionResult);
      
      // Add the range chunks to the cache
      for (const chunk of rangeChunks) {
        this.metadataCache.addChunkWithDependencyAnalysis(chunk);
      }
      
      return rangeChunks;
    } catch (error) {
      console.error(`Error capturing ranges for sheet ${sheet.name}:`, error);
      return [];
    }
  }
}
