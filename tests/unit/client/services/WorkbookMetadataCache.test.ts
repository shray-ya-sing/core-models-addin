import { WorkbookMetadataCache } from "../../../../src/client/services/WorkbookMetadataCache";
import { MetadataChunk } from "../../../../src/client/models/CommandModels";

describe("WorkbookMetadataCache", () => {
  let cache: WorkbookMetadataCache;

  const createChunk = (id: string, type: 'sheet' | 'range'): MetadataChunk => ({
    id,
    type,
    etag: 'e',
    payload: {},
    refs: [],
    lastCaptured: new Date(),
  });

  beforeEach(() => {
    cache = new WorkbookMetadataCache();
  });

  test("setChunk, getChunk, hasChunk", () => {
    const chunk = createChunk('id1', 'sheet');
    expect(cache.getChunk('id1')).toBeNull();
    expect(cache.hasChunk('id1')).toBe(false);

    cache.setChunk(chunk);
    expect(cache.getChunk('id1')).toEqual(chunk);
    expect(cache.hasChunk('id1')).toBe(true);
  });

  test("getAllChunks returns all stored chunks", () => {
    const c1 = createChunk('1', 'sheet');
    const c2 = createChunk('2', 'range');
    cache.setChunk(c1);
    cache.setChunk(c2);
    expect(cache.getAllChunks()).toEqual(expect.arrayContaining([c1, c2]));
  });

  test("getAllSheetChunks filters only sheet chunks", () => {
    const sheetChunk = createChunk('s', 'sheet');
    const rangeChunk = createChunk('r', 'range');
    cache.setChunk(sheetChunk);
    cache.setChunk(rangeChunk);
    expect(cache.getAllSheetChunks()).toEqual([sheetChunk]);
  });

  test("invalidateChunks removes specified and dependent chunks", () => {
    const c1 = createChunk('c1', 'sheet');
    const c2 = createChunk('c2', 'sheet');
    const c3 = createChunk('c3', 'range');
    cache.setChunk(c1);
    cache.setChunk(c2);
    cache.setChunk(c3);
    const analyzer = (cache as any).dependencyAnalyzer;
    analyzer.getTransitiveDependents = jest.fn().mockReturnValue(['c2', 'c3']);

    cache.invalidateChunks(['c1']);
    expect(cache.hasChunk('c1')).toBe(false);
    expect(cache.hasChunk('c2')).toBe(false);
    expect(cache.hasChunk('c3')).toBe(false);
  });

  test("invalidateChunksForSheet invalidates sheet and its ranges", () => {
    const sheetId = 'Sheet:Sheet1';
    const sheetChunk = createChunk(sheetId, 'sheet');
    const rangeId = 'Range:Sheet1!A1:B2';
    const rangeChunk = createChunk(rangeId, 'range');
    const other = createChunk('Other', 'sheet');
    cache.setChunk(sheetChunk);
    cache.setChunk(rangeChunk);
    cache.setChunk(other);

    const analyzer = (cache as any).dependencyAnalyzer;
    analyzer.getTransitiveDependents = jest.fn().mockReturnValue([]);

    cache.invalidateChunksForSheet('Sheet1');
    expect(cache.hasChunk(sheetId)).toBe(false);
    expect(cache.hasChunk(rangeId)).toBe(false);
    expect(cache.hasChunk('Other')).toBe(true);
  });

  test("invalidateAllChunks clears cache and resets dependencies", () => {
    const c1 = createChunk('x', 'sheet');
    cache.setChunk(c1);
    (cache as any).dependencyAnalyzer.resetDependencyGraph = jest.fn();
    cache.invalidateAllChunks();
    expect(cache.getAllChunks()).toHaveLength(0);
    expect((cache as any).dependencyAnalyzer.resetDependencyGraph).toHaveBeenCalled();
  });

  test("get/set workbookVersion", () => {
    expect(cache.getWorkbookVersion()).toBe('');
    cache.setWorkbookVersion('v1');
    expect(cache.getWorkbookVersion()).toBe('v1');
  });
});
