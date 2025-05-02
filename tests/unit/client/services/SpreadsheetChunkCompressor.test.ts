import { SpreadsheetChunkCompressor } from "../../../../src/client/services/SpreadsheetChunkCompressor";
import { SheetState } from "../../../../src/client/models/CommandModels";

describe("SpreadsheetChunkCompressor", () => {
  let compressor: SpreadsheetChunkCompressor;

  beforeEach(() => {
    compressor = new SpreadsheetChunkCompressor();
  });

  describe("processSheetData via compressSheet", () => {
    test("correctly computes metrics and anchors", () => {
      const sheet: SheetState = {
        name: "TestSheet",
        values: [
          [1, 2],
          [3, null]
        ],
        formulas: [
          ["", ""],
          ["=SUM(A1:A1)", ""]
        ],
        tables: [],
        charts: []
      } as any;

      const compressed = compressor.compressSheet(sheet);
      const m = compressed.metrics!;
      expect(m.rowCount).toBe(2);
      expect(m.columnCount).toBe(2);
      expect(m.valueCount).toBe(2);
      expect(m.formulaCount).toBe(1);
      expect(m.emptyCount).toBe(1);
      expect(compressed.anchors).toHaveLength(1);
      expect(compressed.anchors![0]).toMatchObject({ cell: "A2", value: "=SUM(A1:A1)", type: "formula" });
    });

    test("throws error when sheet undefined", () => {
      expect(() => compressor.compressSheet(undefined as any)).toThrow("Sheet is undefined");
    });
  });

  describe("compressSheetToChunk", () => {
    test("produces consistent etag and chunk structure", () => {
      const sheet: SheetState = { name: "CSheet", values: [], formulas: [] } as any;
      const chunk1 = compressor.compressSheetToChunk(sheet);
      const chunk2 = compressor.compressSheetToChunk(sheet);
      expect(chunk1.id).toBe("Sheet:CSheet");
      expect(chunk1.type).toBe("sheet");
      expect(typeof chunk1.etag).toBe("string");
      expect(chunk1.etag).toEqual(chunk2.etag);
      expect(chunk1.payload.summary).toContain("CSheet");
    });
  });

  describe("calculateWorkbookMetrics", () => {
    test("aggregates metrics from multiple chunks", () => {
      const sheet1: SheetState = { name: "S1", values: [["a"]], formulas: [[""]], tables: [], charts: [] } as any;
      const sheet2: SheetState = { name: "S2", values: [["b","c"],["d","" ]], formulas: [["",""],["=SUM(S1!A1)",""]], tables: [], charts: [] } as any;
      const chunk1 = compressor.compressSheetToChunk(sheet1);
      const chunk2 = compressor.compressSheetToChunk(sheet2);
      const metrics = compressor.calculateWorkbookMetrics([chunk1, chunk2]);
      expect(metrics.totalSheets).toBe(2);
      // valueCount from sheet1 =1, sheet2 =3 (b,c,d)
      expect(metrics.totalCells).toBe(chunk1.payload.metrics.valueCount! + chunk2.payload.metrics.valueCount!);
      expect(metrics.totalFormulas).toBe(chunk1.payload.metrics.formulaCount! + chunk2.payload.metrics.formulaCount!);
      expect(metrics.totalTables).toBe(0);
      expect(metrics.totalCharts).toBe(0);
    });
  });

  describe("key detection helpers", () => {
    test("identifies key formulas correctly", () => {
      expect(compressor.test_isKeyFormula("=SUM(A1:A10)")).toBe(true);
      expect(compressor.test_isKeyFormula("=unknown(A1)")).toBe(false);
      expect(compressor.test_isKeyFormula("")).toBe(false);
    });

    test("identifies key values correctly", () => {
      expect(compressor.test_isKeyValue("Total Revenue")).toBe(true);
      expect(compressor.test_isKeyValue("abc")).toBe(false);
      expect(compressor.test_isKeyValue(10000)).toBe(true);
      expect(compressor.test_isKeyValue(50)).toBe(false);
    });
  });
});
