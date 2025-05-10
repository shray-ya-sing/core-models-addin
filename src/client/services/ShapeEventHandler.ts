/**
 * ShapeEventHandler
 * 
 * Handles events for shapes on the Excel drawing layer.
 * This service polls for shape selection events and maps them to actions.
 */

import { PendingChangesTracker } from './pending-changes/PendingChangesTracker';

/**
 * Service for handling shape events in Excel
 */
export class ShapeEventHandler {
  private pendingChangesTracker: PendingChangesTracker;
  private isPolling: boolean = false;
  private pollingInterval: number = 500; // ms
  private refreshInterval: number = 3000; // ms - refresh highlighting every 3 seconds
  private intervalId: number | null = null;
  private refreshIntervalId: number | null = null;
  private lastSelectedShapeName: string | null = null;
  private currentWorkbookId: string = '';
  
  constructor(pendingChangesTracker: PendingChangesTracker) {
    this.pendingChangesTracker = pendingChangesTracker;
  }
  
  /**
   * Set the current workbook ID
   * @param workbookId The ID of the current workbook
   */
  public setCurrentWorkbookId(workbookId: string): void {
    this.currentWorkbookId = workbookId;
    console.log(`ShapeEventHandler: Current workbook ID set to ${workbookId}`);
  }

  /**
   * Start polling for shape selection events and refreshing highlighting
   */
  public startPolling(): void {
    if (this.isPolling) {
      return;
    }
    
    this.isPolling = true;
    
    // Start polling for shape selection events
    this.intervalId = window.setInterval(() => this.checkSelectedShape(), this.pollingInterval);
    
    // Start periodic refresh of highlighting and buttons
    this.refreshIntervalId = window.setInterval(() => this.refreshHighlighting(), this.refreshInterval);
    
    console.log('Started polling for shape selection events and refreshing highlighting');
  }
  
  /**
   * Stop polling for shape selection events and refreshing highlighting
   */
  public stopPolling(): void {
    if (!this.isPolling) {
      return;
    }
    
    // Stop polling for shape selection events
    if (this.intervalId !== null) {
      window.clearInterval(this.intervalId);
      this.intervalId = null;
    }
    
    // Stop periodic refresh of highlighting and buttons
    if (this.refreshIntervalId !== null) {
      window.clearInterval(this.refreshIntervalId);
      this.refreshIntervalId = null;
    }
    
    this.isPolling = false;
    console.log('Stopped polling for shape selection events and refreshing highlighting');
  }
  
  /**
   * Refresh highlighting and buttons for all pending changes
   */
  private async refreshHighlighting(): Promise<void> {
    if (!this.currentWorkbookId) {
      return;
    }
    
    await this.pendingChangesTracker.refreshPendingChangesHighlighting(this.currentWorkbookId);
  }
  
  /**
   * Check if a shape is selected and handle the event
   */
  private async checkSelectedShape(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        // Get the active worksheet and its shapes
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = worksheet.shapes;
        
        // Get the selected shape using the getSelection API
        const selection = context.workbook.getSelectedRange();
        selection.load('address');
        shapes.load(['items', 'count']);
        
        await context.sync();
        
        // Check if we have shapes to process
        if (shapes.items.length === 0) {
          return;
        }
        
        // Process all shapes to check for clicks
        for (const shape of shapes.items) {
          shape.load(['name', 'left', 'top', 'width', 'height']);
        }
        
        await context.sync();
        
        // Find shapes that might have been clicked
        // We'll use a simple approach: check if any shape's name starts with 'accept-' or 'reject-'
        // and if it's different from the last processed shape
        for (const shape of shapes.items) {
          const shapeName = shape.name;
          
          // Check if this is an accept/reject button and hasn't been processed yet
          if ((shapeName.startsWith('accept-') || shapeName.startsWith('reject-')) && 
              shapeName !== this.lastSelectedShapeName) {
            
            // Process this shape
            this.lastSelectedShapeName = shapeName;
            
            // Handle the shape click
            await this.handleShapeClick(shapeName);
            
            // Break after processing one shape to avoid multiple triggers
            break;
          }
        }
        
        await context.sync();
      });
    } catch (error) {
      // Ignore errors during polling
      if (error instanceof Error && error.message.indexOf('ItemNotFound') === -1) {
        console.error('Error checking selected shape:', error);
      }
    }
  }
  
  /**
   * Handle a shape click event
   * @param shapeName The name of the clicked shape
   */
  private async handleShapeClick(shapeName: string): Promise<void> {
    // Check if this is an accept button
    if (shapeName.startsWith('accept-')) {
      const changeId = shapeName.substring(7); // Remove 'accept-' prefix
      await this.pendingChangesTracker.acceptChange(changeId);
      console.log(`Accept button clicked for change: ${changeId}`);
    }
    
    // Check if this is a reject button
    else if (shapeName.startsWith('reject-')) {
      const changeId = shapeName.substring(7); // Remove 'reject-' prefix
      await this.pendingChangesTracker.rejectChange(changeId);
      console.log(`Reject button clicked for change: ${changeId}`);
    }
  }
}
