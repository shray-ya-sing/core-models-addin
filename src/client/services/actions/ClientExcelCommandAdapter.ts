// src/client/services/ClientExcelCommandAdapter.ts
// Adapter that connects the Command system to the Excel Operations DSL

import { Command, Operation, OperationType } from '../../models/CommandModels';
import { ExcelOperation } from '../../models/ExcelOperationModels';
import { ClientExcelCommandInterpreter } from './ClientExcelCommandInterpreter';

/**
 * Adapter that connects the Command system to the Excel Operations DSL
 * This class bridges the gap between the existing command system and the new Excel operations DSL
 */
export class ClientExcelCommandAdapter {
  private interpreter: ClientExcelCommandInterpreter;

  /**
   * Create a new ClientExcelCommandAdapter
   * @param interpreter Optional existing interpreter instance to use. If not provided, a new one will be created.
   */
  constructor(interpreter?: ClientExcelCommandInterpreter) {
    this.interpreter = interpreter || new ClientExcelCommandInterpreter();
    console.log(`🔄 [ClientExcelCommandAdapter] Using ${interpreter ? 'provided' : 'new'} interpreter instance`);
  }

  /**
   * Execute a command using the Excel Operations DSL
   * @param command The command to execute
   * @returns Array of operation types that were executed
   */
  public async executeCommand(command: Command): Promise<string[]> {
    try {
      console.log(`🔎 [ClientExcelCommandAdapter] executeCommand called for command: ${command.id} - ${command.description}`);
      
      // Track all operation types that were executed
      const executedOperationTypes: string[] = [];
      
      // Track operations by ID to detect duplicates
      const operationIds = new Set<string>();
      const operationFingerprints = new Map<string, number>(); // Map of fingerprint to count
      
      // Process each step in the command
      for (const step of command.steps) {
        // Get step index from the steps array
        const stepIndex = command.steps.indexOf(step);
        console.log(`📑 [ClientExcelCommandAdapter] Processing step ${stepIndex}: ${step.description}`);
        
        // Extract Excel operations from the step operations
        const excelOperations = this.extractExcelOperations(step.operations);
        console.log(`📊 [ClientExcelCommandAdapter] Extracted ${excelOperations.length} Excel operations from step`);
        
        if (excelOperations.length > 0) {
          // Check for duplicate operations
          excelOperations.forEach(op => {
            // Create a fingerprint for the operation
            const fingerprint = this.createOperationFingerprint(op);
            
            // Check if we've seen this operation ID before
            if (op.id && operationIds.has(op.id)) {
              console.warn(`⚠️ [ClientExcelCommandAdapter] Duplicate operation ID detected: ${op.id}`);
            } else if (op.id) {
              operationIds.add(op.id);
            }
            
            // Check if we've seen this operation fingerprint before
            if (operationFingerprints.has(fingerprint)) {
              const count = operationFingerprints.get(fingerprint) || 0;
              operationFingerprints.set(fingerprint, count + 1);
              console.warn(`⚠️ [ClientExcelCommandAdapter] Duplicate operation detected by fingerprint: ${fingerprint} (count: ${count + 1})`);
            } else {
              operationFingerprints.set(fingerprint, 1);
            }
            
            // Collect operation types for cache invalidation
            if (op.op && !executedOperationTypes.includes(op.op)) {
              executedOperationTypes.push(op.op);
            }
          });
          
          console.log(`🔄 [ClientExcelCommandAdapter] Executing ${excelOperations.length} operations`);
          // Execute the Excel operations using the interpreter
          await this.interpreter.executeOperations(excelOperations);
        }
      }
      
      // Log summary of duplicate operations
      let duplicateCount = 0;
      operationFingerprints.forEach((count, fingerprint) => {
        if (count > 1) {
          duplicateCount++;
          console.warn(`🔍 [ClientExcelCommandAdapter] Found ${count} instances of operation: ${fingerprint}`);
        }
      });
      
      if (duplicateCount > 0) {
        console.warn(`🚨 [ClientExcelCommandAdapter] Found ${duplicateCount} duplicate operations in command`);
      } else {
        console.log(`✅ [ClientExcelCommandAdapter] No duplicate operations detected in command`);
      }
      
      // Return the operation types that were executed
      return executedOperationTypes;
    } catch (error) {
      console.error('Error executing Excel command:', error);
      throw error;
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
    }
    
    // Create a fingerprint string that combines these properties
    const fingerprint = `${opType}:${targetOrRange}:${valueOrFormula}:${formatInfo}`;
    
    return fingerprint;
  }
  
  /**
   * Extract Excel operations from command operations
   * @param operations The command operations
   * @returns The Excel operations
   */
  private extractExcelOperations(operations: Operation[]): ExcelOperation[] {
    const excelOperations: ExcelOperation[] = [];

    for (const operation of operations) {
      // Check if the operation value is an Excel operation
      if (operation.value && typeof operation.value === 'object' && 'op' in operation.value) {
        excelOperations.push(operation.value as ExcelOperation);
      }
    }

    return excelOperations;
  }
}
