// src/client/services/ClientExcelCommandAdapter.ts
// Adapter that connects the Command system to the Excel Operations DSL

import { Command, Operation, OperationType } from '../models/CommandModels';
import { ExcelOperation } from '../models/ExcelOperationModels';
import { ClientExcelCommandInterpreter } from './ClientExcelCommandInterpreter';

/**
 * Adapter that connects the Command system to the Excel Operations DSL
 * This class bridges the gap between the existing command system and the new Excel operations DSL
 */
export class ClientExcelCommandAdapter {
  private interpreter: ClientExcelCommandInterpreter;

  constructor() {
    this.interpreter = new ClientExcelCommandInterpreter();
  }

  /**
   * Execute a command using the Excel Operations DSL
   * @param command The command to execute
   */
  public async executeCommand(command: Command): Promise<void> {
    try {
      // Process each step in the command
      for (const step of command.steps) {
        // Extract Excel operations from the step operations
        const excelOperations = this.extractExcelOperations(step.operations);
        
        if (excelOperations.length > 0) {
          // Execute the Excel operations using the interpreter
          await this.interpreter.executeOperations(excelOperations);
        }
      }
    } catch (error) {
      console.error('Error executing Excel command:', error);
      throw error;
    }
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
