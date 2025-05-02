import { EmbeddingStore, EmbeddingVector } from "../../../../src/client/services/EmbeddingStore";
describe("EmbeddingStore", () => {
  let store: EmbeddingStore;
  const sheetChunk = { id: 'c1', etag: 'e1', type: 'sheet', payload: { name: 'Sheet1', summary: 'Summary', anchors: [{ value: 'v1', address: 'A1' }], values: [{ value: 10, address: 'B2' }] } };
  const rangeChunk = { id: 'r1', etag: 'e2', type: 'range', payload: { sheet: 'Sheet1', address: 'A1:B2', description: 'Desc', values: [[1, 2], [3, 4]] } };

  beforeEach(() => {
    store = new EmbeddingStore({ persistToLocalStorage: false, useLocalModel: true });
  });

  test("initialize sets isInitialized flag only once", async () => {
    expect((store as any).isInitialized).toBe(false);
    await store.initialize();
    expect((store as any).isInitialized).toBe(true);
    await store.initialize();
    expect((store as any).isInitialized).toBe(true);
  });

  test("computeSimpleEmbedding produces 128-dim unit vector", () => {
    const vec = (store as any).computeSimpleEmbedding("abc");
    expect(vec).toHaveLength(128);
    const mag = vec.reduce((sum: number, v: number) => sum + v * v, 0);
    expect(mag).toBeCloseTo(1, 5);
  });

  test("getTextRepresentation handles sheet and range chunks", () => {
    const sheetText = (store as any).getTextRepresentation(sheetChunk);
    expect(sheetText).toContain("Sheet: Sheet1");
    expect(sheetText).toContain("Summary: Summary");
    expect(sheetText).toContain("Key cells:");

    const rangeText = (store as any).getTextRepresentation(rangeChunk);
    expect(rangeText).toContain("Range: Sheet1!A1:B2");
    expect(rangeText).toContain("Description: Desc");
    expect(rangeText).toContain("Values:");
  });

  test("getEmbedding computes and caches embedding", async () => {
    const v1 = await store.getEmbedding(sheetChunk as any);
    expect(v1).toHaveLength(128);
    const cached = (store as any).embeddings.get('c1').vector;
    expect(cached).toBe(v1);
    const v2 = await store.getEmbedding(sheetChunk as any);
    expect(v2).toBe(v1);
  });

  test("getEmbedding with forceRefresh updates timestamp", async () => {
    const v1 = await store.getEmbedding(sheetChunk as any);
    const meta1 = (store as any).embeddings.get('c1').metadata;
    const ts1 = meta1.createdAt;
    await new Promise((r) => setTimeout(r, 10));
    const v2 = await store.getEmbedding(sheetChunk as any, true);
    const meta2 = (store as any).embeddings.get('c1').metadata;
    expect(meta2.createdAt).toBeGreaterThan(ts1);
    expect(v2).toEqual(v1);
  });

  test("findSimilarChunks returns sorted results", async () => {
    const chunkA = { id: 'a', etag: 'eA', type: 'sheet', payload: { name: 'Alpha', summary: '', anchors: [], values: [] } };
    const chunkB = { id: 'b', etag: 'eB', type: 'sheet', payload: { name: 'Beta', summary: '', anchors: [], values: [] } };
    const results = await store.findSimilarChunks('Alpha', [chunkA as any, chunkB as any], 2);
    expect(results).toHaveLength(2);
    expect(results[0].chunkId).toBe('a');
    expect(results[1].chunkId).toBe('b');
  });
});
