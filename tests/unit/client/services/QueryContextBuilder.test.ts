import { QueryContextBuilder } from "../../../../src/client/services/QueryContextBuilder";
import { QueryType } from "../../../../src/client/models/CommandModels";

describe("QueryContextBuilder", () => {
  let builder: QueryContextBuilder;
  let workbookStateManager: any;
  let metadataCache: any;
  let chunkLocator: any;

  beforeEach(() => {
    workbookStateManager = {
      getActiveSheetName: jest.fn().mockResolvedValue("ActiveSheet"),
      captureWorkbookState: jest.fn(),
      getChunkCompressor: jest.fn()
    };
    metadataCache = {
      getAllSheetChunks: jest.fn().mockReturnValue(['dummy']),
      getRelatedChunks: jest.fn(),
      calculateWorkbookMetrics: jest.fn(),
      getAllChunks: jest.fn()
    };
    chunkLocator = {
      setActiveSheet: jest.fn(),
      locateChunks: jest.fn()
    };
    builder = new QueryContextBuilder(
      workbookStateManager,
      metadataCache,
      chunkLocator
    );
  });

  test("setChunkLocator sets the new locator", () => {
    const newLocator = { foo: "bar" };
    builder.setChunkLocator(newLocator as any);
    expect((builder as any).chunkLocator).toBe(newLocator);
  });

  test("buildContextForQuery fallback includes all sheets when locator returns empty", async () => {
    const sheetChunk = { id: "Sheet1", type: 'sheet', payload: { name: "Sheet1" } };
    metadataCache.getAllSheetChunks.mockReturnValue([sheetChunk]);
    chunkLocator.locateChunks.mockResolvedValue({ chunkIds: [], details: { sheets: [] }, confidenceScores: new Map(), usedLLM: false });
    metadataCache.calculateWorkbookMetrics.mockReturnValue({ totalSheets: 1, totalCells: 0, totalFormulas: 0, totalTables: 0, totalCharts: 0 });
    metadataCache.getAllChunks.mockReturnValue([sheetChunk]);

    const context = await builder.buildContextForQuery(QueryType.WorkbookQuestion, [], "any");
    expect(context.chunks).toEqual([sheetChunk]);
    expect(context.activeSheet).toBe("ActiveSheet");
    expect(context.metrics).toEqual({ totalSheets: 1, totalCells: 0, totalFormulas: 0, totalTables: 0, totalCharts: 0 });
  });

  test("buildContextForQuery uses locator results", async () => {
    const relatedChunks = [{ id: "x" }];
    metadataCache.getAllSheetChunks.mockReturnValue(['Sheet:Dummy']);
    chunkLocator.locateChunks.mockResolvedValue({ chunkIds: ["x"], details: { sheets: ["S1"] }, confidenceScores: new Map(), usedLLM: false });
    metadataCache.getRelatedChunks.mockReturnValue(relatedChunks);
    metadataCache.calculateWorkbookMetrics.mockReturnValue({ totalSheets: 0, totalCells: 0, totalFormulas: 0, totalTables: 0, totalCharts: 0 });

    const context = await builder.buildContextForQuery(QueryType.WorkbookQuestion, [], "q");
    expect(metadataCache.getRelatedChunks).toHaveBeenCalledWith(["x"]);
    expect(context.chunks).toBe(relatedChunks);
    expect(context.activeSheet).toBe("ActiveSheet");
  });

  test("contextToJson outputs correct structure", () => {
    const sheetPayload = { name: "S", summary: "sum", anchors: [], values: [] };
    const context: any = {
      chunks: [{ type: 'sheet', payload: sheetPayload }],
      activeSheet: "S",
      metrics: { totalSheets: 1, totalCells: 2, totalFormulas: 3, totalTables: 4, totalCharts: 5 }
    };

    const json = builder.contextToJson(context);
    const obj = JSON.parse(json);
    expect(obj.sheets).toEqual([sheetPayload]);
    expect(obj.activeSheet).toBe("S");
    expect(obj.metrics).toEqual(context.metrics);
    expect(obj._diagnostic.chunkCount).toBe(1);
    expect(obj._diagnostic.validSheetCount).toBe(1);
  });
});
