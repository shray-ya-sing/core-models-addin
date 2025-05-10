/**
 * Analyzer for formula dependencies between sheets and ranges
 * Used to build a dependency graph for targeted workbook state capture
 */
import { MetadataChunk } from '../../models/CommandModels';

/**
 * Represents a dependency graph between different chunks of workbook data
 */
export interface DependencyGraph {
  // Map from chunk ID to set of chunk IDs it depends on
  forward: Map<string, Set<string>>;
  // Map from chunk ID to set of chunk IDs that depend on it
  reverse: Map<string, Set<string>>;
}

/**
 * Analyzes and tracks dependencies between sheets and ranges
 * in an Excel workbook to enable selective state capture
 */
export class RangeDependencyAnalyzer {
  private dependencyGraph: DependencyGraph = {
    forward: new Map<string, Set<string>>(),
    reverse: new Map<string, Set<string>>()
  };
  
  constructor() {
    this.resetDependencyGraph();
  }
  
  /**
   * Reset the dependency graph to an empty state
   */
  public resetDependencyGraph(): void {
    this.dependencyGraph = {
      forward: new Map<string, Set<string>>(),
      reverse: new Map<string, Set<string>>()
    };
  }
  
  /**
   * Get the current dependency graph
   * @returns The dependency graph
   */
  public getDependencyGraph(): DependencyGraph {
    return this.dependencyGraph;
  }
  
  /**
   * Add a dependency from one chunk to another
   * @param sourceId The ID of the source chunk (depends on targetId)
   * @param targetId The ID of the target chunk (is depended on by sourceId)
   */
  public addDependency(sourceId: string, targetId: string): void {
    // Skip self-dependencies
    if (sourceId === targetId) {
      return;
    }
    
    // Add to forward dependencies (source -> targets it depends on)
    if (!this.dependencyGraph.forward.has(sourceId)) {
      this.dependencyGraph.forward.set(sourceId, new Set<string>());
    }
    this.dependencyGraph.forward.get(sourceId)!.add(targetId);
    
    // Add to reverse dependencies (target -> sources that depend on it)
    if (!this.dependencyGraph.reverse.has(targetId)) {
      this.dependencyGraph.reverse.set(targetId, new Set<string>());
    }
    this.dependencyGraph.reverse.get(targetId)!.add(sourceId);
  }
  
  /**
   * Remove a dependency
   * @param sourceId The ID of the source chunk
   * @param targetId The ID of the target chunk
   */
  public removeDependency(sourceId: string, targetId: string): void {
    // Remove from forward dependencies
    if (this.dependencyGraph.forward.has(sourceId)) {
      this.dependencyGraph.forward.get(sourceId)!.delete(targetId);
    }
    
    // Remove from reverse dependencies
    if (this.dependencyGraph.reverse.has(targetId)) {
      this.dependencyGraph.reverse.get(targetId)!.delete(sourceId);
    }
  }
  
  /**
   * Remove all dependencies for a chunk
   * @param chunkId The chunk ID to remove dependencies for
   */
  public removeAllDependenciesForChunk(chunkId: string): void {
    // Remove forward dependencies
    if (this.dependencyGraph.forward.has(chunkId)) {
      const targets = Array.from(this.dependencyGraph.forward.get(chunkId)!);
      
      // For each target, remove this chunk from its reverse dependencies
      for (const targetId of targets) {
        if (this.dependencyGraph.reverse.has(targetId)) {
          this.dependencyGraph.reverse.get(targetId)!.delete(chunkId);
        }
      }
      
      // Remove the chunk's forward dependencies
      this.dependencyGraph.forward.delete(chunkId);
    }
    
    // Remove reverse dependencies
    if (this.dependencyGraph.reverse.has(chunkId)) {
      const sources = Array.from(this.dependencyGraph.reverse.get(chunkId)!);
      
      // For each source, remove this chunk from its forward dependencies
      for (const sourceId of sources) {
        if (this.dependencyGraph.forward.has(sourceId)) {
          this.dependencyGraph.forward.get(sourceId)!.delete(chunkId);
        }
      }
      
      // Remove the chunk's reverse dependencies
      this.dependencyGraph.reverse.delete(chunkId);
    }
  }
  
  /**
   * Get all chunks that depend on the given chunk
   * @param chunkId The chunk ID to find dependents for
   * @returns Set of chunk IDs that depend on the given chunk
   */
  public getDependentChunks(chunkId: string): Set<string> {
    if (this.dependencyGraph.reverse.has(chunkId)) {
      return new Set(this.dependencyGraph.reverse.get(chunkId)!);
    }
    return new Set<string>();
  }
  
  /**
   * Get all chunks that the given chunk depends on
   * @param chunkId The chunk ID to find dependencies for
   * @returns Set of chunk IDs that the given chunk depends on
   */
  public getDependencyChunks(chunkId: string): Set<string> {
    if (this.dependencyGraph.forward.has(chunkId)) {
      return new Set(this.dependencyGraph.forward.get(chunkId)!);
    }
    return new Set<string>();
  }
  
  /**
   * Get the transitive closure of dependencies
   * Returns all chunks that are directly or indirectly dependent on the given chunks
   * @param chunkIds The starting chunk IDs
   * @returns Set of all dependent chunk IDs
   */
  public getTransitiveDependents(chunkIds: string[]): Set<string> {
    const result = new Set<string>();
    const queue = [...chunkIds];
    
    while (queue.length > 0) {
      const current = queue.shift()!;
      
      // Skip if already processed
      if (result.has(current)) {
        continue;
      }
      
      // Add to result
      result.add(current);
      
      // Add dependents to queue
      const dependents = this.getDependentChunks(current);
      for (const dependent of dependents) {
        if (!result.has(dependent)) {
          queue.push(dependent);
        }
      }
    }
    
    // Remove the original chunks from the result
    for (const chunkId of chunkIds) {
      result.delete(chunkId);
    }
    
    return result;
  }
  
  /**
   * Get the transitive closure of dependencies
   * Returns all chunks that the given chunks directly or indirectly depend on
   * @param chunkIds The starting chunk IDs
   * @returns Set of all dependency chunk IDs
   */
  public getTransitiveDependencies(chunkIds: string[]): Set<string> {
    const result = new Set<string>();
    const queue = [...chunkIds];
    
    while (queue.length > 0) {
      const current = queue.shift()!;
      
      // Skip if already processed
      if (result.has(current)) {
        continue;
      }
      
      // Add to result
      result.add(current);
      
      // Add dependencies to queue
      const dependencies = this.getDependencyChunks(current);
      for (const dependency of dependencies) {
        if (!result.has(dependency)) {
          queue.push(dependency);
        }
      }
    }
    
    // Remove the original chunks from the result
    for (const chunkId of chunkIds) {
      result.delete(chunkId);
    }
    
    return result;
  }
  
  /**
   * Get all chunks related to the given chunks
   * Includes both dependencies and dependents
   * @param chunkIds The starting chunk IDs
   * @returns Set of all related chunk IDs
   */
  public getAllRelatedChunks(chunkIds: string[]): Set<string> {
    const dependencies = this.getTransitiveDependencies(chunkIds);
    const dependents = this.getTransitiveDependents(chunkIds);
    
    // Combine both sets
    return new Set([...dependencies, ...dependents]);
  }
  
  /**
   * Analyze a set of chunks and build the dependency graph
   * @param chunks The chunks to analyze
   */
  public analyzeChunks(chunks: MetadataChunk[]): void {
    // First, clear existing dependencies for these chunks
    for (const chunk of chunks) {
      this.removeAllDependenciesForChunk(chunk.id);
    }
    
    // Then build new dependencies based on the refs property
    for (const chunk of chunks) {
      for (const refId of chunk.refs) {
        this.addDependency(chunk.id, refId);
      }
    }
  }
  
  /**
   * Extract dependencies from a formula
   * @param formula The formula to extract dependencies from
   * @param sourceName The name of the sheet containing the formula
   * @returns Array of sheet names referenced in the formula
   */
  public extractSheetReferencesFromFormula(formula: string, sourceName: string): string[] {
    if (!formula || typeof formula !== 'string') {
      return [];
    }
    
    // Remove any leading spaces and check if it's a formula
    const trimmedFormula = formula.trim();
    if (!trimmedFormula.startsWith('=')) {
      return [];
    }
    
    const sheetRefs = new Set<string>();
    
    try {
      // Pattern 1: Standard sheet references like Sheet1!A1 or Sheet1!A1:B10
      // Also handles quotes for sheet names with spaces like 'Income Statement'!A1
      const standardRefRegex = /([\w\s']+)!([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)/g;
      let match;
      
      while ((match = standardRefRegex.exec(trimmedFormula)) !== null) {
        const sheetName = match[1].replace(/'/g, ''); // Remove any single quotes
        
        // Skip self-references
        if (sheetName !== sourceName) {
          sheetRefs.add(sheetName);
        }
      }
      
      // Pattern 2: INDIRECT references like INDIRECT("Sheet1!A1")
      const indirectRefRegex = /INDIRECT\(\s*"([\w\s]+)!([A-Z0-9:]+)"\s*\)/gi;
      
      while ((match = indirectRefRegex.exec(trimmedFormula)) !== null) {
        const sheetName = match[1];
        
        if (sheetName !== sourceName) {
          sheetRefs.add(sheetName);
        }
      }
      
      // Pattern 3: Sheet names in 3D references like SUM(Sheet1:Sheet3!A1)
      const threeDRefRegex = /([\w\s']+):([\w\s']+)!/g;
      
      while ((match = threeDRefRegex.exec(trimmedFormula)) !== null) {
        const startSheet = match[1].replace(/'/g, '');
        const endSheet = match[2].replace(/'/g, '');
        
        if (startSheet !== sourceName) {
          sheetRefs.add(startSheet);
        }
        
        if (endSheet !== sourceName) {
          sheetRefs.add(endSheet);
        }
      }
      
      // Pattern 4: CELL, ADDRESS and other functions that might contain sheet names
      const cellRefRegex = /CELL\(\s*"[^"]*"\s*,\s*'?([\w\s]+)'?!([A-Z0-9:]+)\s*\)/gi;
      
      while ((match = cellRefRegex.exec(trimmedFormula)) !== null) {
        const sheetName = match[1];
        
        if (sheetName !== sourceName) {
          sheetRefs.add(sheetName);
        }
      }
    } catch (error) {
      console.warn(`Error extracting sheet references from formula: ${formula}`, error);
    }
    
    return Array.from(sheetRefs);
  }
  
  /**
   * Analyze formulas in a sheet and update the dependency graph
   * @param sheetName The name of the sheet
   * @param formulas The 2D array of formulas in the sheet
   */
  public analyzeFormulasInSheet(sheetName: string, formulas: any[][]): void {
    if (!formulas || !Array.isArray(formulas)) {
      return;
    }
    
    const sheetId = `Sheet:${sheetName}`;
    const referencedSheets = new Set<string>();
    
    // Process each formula in the sheet
    for (const row of formulas) {
      if (!Array.isArray(row)) continue;
      
      for (const cell of row) {
        if (!cell || typeof cell !== 'string' || !cell.startsWith('=')) continue;
        
        // Extract sheet references from the formula
        const sheetRefs = this.extractSheetReferencesFromFormula(cell, sheetName);
        
        // Add each reference to the set
        for (const ref of sheetRefs) {
          referencedSheets.add(`Sheet:${ref}`);
        }
      }
    }
    
    // Add dependencies to the graph
    for (const refSheetId of referencedSheets) {
      this.addDependency(sheetId, refSheetId);
    }
  }
}
