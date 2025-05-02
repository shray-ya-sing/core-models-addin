import { ChunkLocatorService } from "../../../../src/client/services/ChunkLocatorService";
import { WorkbookMetadataCache } from "../../../../src/client/services/WorkbookMetadataCache";
import { RangeDependencyAnalyzer } from "../../../../src/client/services/RangeDependencyAnalyzer";

describe("ChunkLocatorService", () => {
  let service: ChunkLocatorService;
  let metadataCache: jest.Mocked<WorkbookMetadataCache>;
  let embeddingStore: any;
  let dependencyAnalyzer: RangeDependencyAnalyzer;

  beforeEach(() => {
    metadataCache = {
      getMetadata: jest.fn(),
      getAllSheetChunks: jest.fn().mockReturnValue([])
    } as any;
    embeddingStore = {
      initialize: jest.fn().mockResolvedValue(undefined),
      getEmbedding: jest.fn(),
      findSimilarChunks: jest.fn(),
      clear: jest.fn()
    } as any;
    dependencyAnalyzer = new RangeDependencyAnalyzer();

    service = new ChunkLocatorService({
      metadataCache,
      embeddingStore,
      dependencyAnalyzer,
      config: { enableLLM: false, useNaiveLLMSelection: false }
    });
  });

  test("setActiveSheet sets the active sheet name", () => {
    expect((service as any).activeSheetName).toBeNull();
    service.setActiveSheet("Sheet1");
    expect((service as any).activeSheetName).toBe("Sheet1");
  });

  test("locateChunks returns default result when no matches", async () => {
    const result = await service.locateChunks("any query", []);
    expect(result.chunkIds).toEqual([]);
    expect(result.details.sheets).toEqual([]);
    expect(result.details.ranges).toEqual([]);
    expect(result.details.charts).toEqual([]);
    expect(result.confidenceScores.size).toBe(0);
    expect(result.usedLLM).toBe(false);
  });
});
