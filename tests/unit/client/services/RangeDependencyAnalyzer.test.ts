/**
 * Unit tests for RangeDependencyAnalyzer
 */

import { RangeDependencyAnalyzer } from "../../../../src/client/services/RangeDependencyAnalyzer";
import { MetadataChunk } from "../../../../src/client/models/CommandModels";

describe("RangeDependencyAnalyzer", () => {
  let analyzer: RangeDependencyAnalyzer;
  
  beforeEach(() => {
    analyzer = new RangeDependencyAnalyzer();
  });
  
  describe("dependency graph operations", () => {
    test("adds and retrieves dependencies correctly", () => {
      // Add dependencies
      analyzer.addDependency("Sheet1!A1", "Sheet1!B1");
      analyzer.addDependency("Sheet1!A1", "Sheet1!C1");
      analyzer.addDependency("Sheet1!C1", "Sheet1!D1");
      
      // Verify forward dependencies
      const aDependencies = analyzer.getDependencyChunks("Sheet1!A1");
      expect(aDependencies.has("Sheet1!B1")).toBe(true);
      expect(aDependencies.has("Sheet1!C1")).toBe(true);
      expect(aDependencies.size).toBe(2);
      
      // Verify reverse dependencies
      const bDependents = analyzer.getDependentChunks("Sheet1!B1");
      expect(bDependents.has("Sheet1!A1")).toBe(true);
      expect(bDependents.size).toBe(1);
    });
    
    test("removes dependencies correctly", () => {
      // Setup dependencies
      analyzer.addDependency("Sheet1!A1", "Sheet1!B1");
      analyzer.addDependency("Sheet1!A1", "Sheet1!C1");
      
      // Remove one dependency
      analyzer.removeDependency("Sheet1!A1", "Sheet1!B1");
      
      // Verify the dependency was removed
      const dependencies = analyzer.getDependencyChunks("Sheet1!A1");
      expect(dependencies.has("Sheet1!B1")).toBe(false);
      expect(dependencies.has("Sheet1!C1")).toBe(true);
      expect(dependencies.size).toBe(1);
    });
    
    test("removes all dependencies for a chunk", () => {
      // Setup dependencies
      analyzer.addDependency("Sheet1!A1", "Sheet1!B1");
      analyzer.addDependency("Sheet1!A1", "Sheet1!C1");
      analyzer.addDependency("Sheet1!D1", "Sheet1!A1");
      
      // Remove all dependencies for A1
      analyzer.removeAllDependenciesForChunk("Sheet1!A1");
      
      // Verify all dependencies are removed
      expect(analyzer.getDependencyChunks("Sheet1!A1").size).toBe(0);
      expect(analyzer.getDependentChunks("Sheet1!A1").size).toBe(0);
      
      // Verify other dependencies remain intact
      expect(analyzer.getDependentChunks("Sheet1!B1").size).toBe(0);
      expect(analyzer.getDependentChunks("Sheet1!C1").size).toBe(0);
    });
  });
  
  describe("transitive dependency analysis", () => {
    beforeEach(() => {
      // Setup a dependency chain for tests
      analyzer.addDependency("Sheet1!A1", "Sheet1!B1");
      analyzer.addDependency("Sheet1!B1", "Sheet1!C1");
      analyzer.addDependency("Sheet1!C1", "Sheet1!D1");
      analyzer.addDependency("Sheet1!X1", "Sheet1!Y1");
    });
    
    test("finds transitive dependencies", () => {
      const dependencies = analyzer.getTransitiveDependencies(["Sheet1!A1"]);
      
      expect(dependencies.has("Sheet1!B1")).toBe(true);
      expect(dependencies.has("Sheet1!C1")).toBe(true);
      expect(dependencies.has("Sheet1!D1")).toBe(true);
      expect(dependencies.has("Sheet1!Y1")).toBe(false); // Not related to A1
    });
    
    test("finds transitive dependents", () => {
      const dependents = analyzer.getTransitiveDependents(["Sheet1!D1"]);
      
      expect(dependents.has("Sheet1!C1")).toBe(true);
      expect(dependents.has("Sheet1!B1")).toBe(true);
      expect(dependents.has("Sheet1!A1")).toBe(true);
      expect(dependents.has("Sheet1!X1")).toBe(false); // Not related to D1
    });
    
    test("finds all related chunks", () => {
      const related = analyzer.getAllRelatedChunks(["Sheet1!B1"]);
      
      expect(related.has("Sheet1!A1")).toBe(true); // Dependent
      expect(related.has("Sheet1!C1")).toBe(true); // Dependency
      expect(related.has("Sheet1!D1")).toBe(true); // Transitive dependency
      expect(related.has("Sheet1!X1")).toBe(false); // Not related
      expect(related.has("Sheet1!Y1")).toBe(false); // Not related
    });
  });
  
  describe("analyzeChunks", () => {
    test("analyzes chunk dependencies correctly", () => {
      const chunks: MetadataChunk[] = [
        {
          id: "chunk1",
          type: "sheet",
          etag: "hash1",
          payload: { name: "Sheet1" },
          refs: ["chunk2", "chunk3"],
          lastCaptured: new Date()
        },
        {
          id: "chunk2",
          type: "range",
          etag: "hash2",
          payload: { address: "A1:B10", sheet: "Sheet1" },
          refs: ["chunk4"],
          lastCaptured: new Date()
        },
        {
          id: "chunk3",
          type: "range",
          etag: "hash3",
          payload: { address: "C1:D10", sheet: "Sheet1" },
          refs: [],
          lastCaptured: new Date()
        },
        {
          id: "chunk4",
          type: "range",
          etag: "hash4",
          payload: { address: "A1:A5", sheet: "Sheet2" },
          refs: [],
          lastCaptured: new Date()
        }
      ];
      
      analyzer.analyzeChunks(chunks);
      
      // Verify dependencies were established correctly
      expect(analyzer.getDependencyChunks("chunk1").has("chunk2")).toBe(true);
      expect(analyzer.getDependencyChunks("chunk1").has("chunk3")).toBe(true);
      expect(analyzer.getDependencyChunks("chunk2").has("chunk4")).toBe(true);
      
      // Verify transitive dependencies
      expect(analyzer.getTransitiveDependencies(["chunk1"]).has("chunk4")).toBe(true);
    });
  });
});

