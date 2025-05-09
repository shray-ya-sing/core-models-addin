import { RangeDetector } from "../../../../src/client/services/RangeDetector";
import { SheetState } from "../../../../src/client/models/CommandModels";
import { RangeType } from "../../../../src/client/models/RangeModels";
import { createRangeChunk, formatRangeId } from "../../../../src/client/models/RangeModels";

describe("RangeDetector", () => {
  let detector: RangeDetector;

  beforeEach(() => {
    detector = new RangeDetector();
  });

  test("returns empty for null or invalid sheet", () => {
    const result = detector.detectRanges(null as any);
    expect(result.ranges).toEqual([]);
    expect(result.rangeIdToChunkId.size).toBe(0);
  });

  test("detects tables in sheet state", () => {
    const sheet: SheetState = {
      name: "Sheet1",
      values: [[1, 2], [3, 4]],
      formulas: [],
      tables: [
        { range: "A1:B2", name: "Table1", headers: ["H1", "H2"] }
      ],
      namedRanges: []
    } as any;

    const result = detector.detectRanges(sheet);
    expect(result.ranges).toHaveLength(1);
    const r = result.ranges[0];
    expect(r.type).toBe(RangeType.DataTable);
    expect(r.range).toBe("A1:B2");
    expect(r.name).toBe("Table1");
    expect(r.rowCount).toBe(2);
    expect(r.columnCount).toBe(2);
    expect(result.rangeIdToChunkId.get(r.range)).toBe(formatRangeId(r));
  });

  test("detects named ranges in sheet state", () => {
    const sheet: SheetState = {
      name: "Sheet2",
      values: [["x", "y", "z"]],
      formulas: [],
      tables: [],
      namedRanges: [
        { value: "Sheet2!A1:C1", name: "NR1" }
      ]
    } as any;

    const result = detector.detectRanges(sheet);
    expect(result.ranges).toHaveLength(1);
    const r = result.ranges[0];
    expect(r.type).toBe(RangeType.NamedRange);
    expect(r.range).toBe("A1:C1");
    expect(r.name).toBe("NR1");
    expect(r.rowCount).toBe(1);
    expect(r.columnCount).toBe(3);
  });

  test("createRangeChunks produces MetadataChunks", () => {
    const sheet: SheetState = {
      name: "Sheet3",
      values: [[10]],
      formulas: [["=A1"]],
      tables: [],
      namedRanges: []
    } as any;
    const detection = {
      ranges: [
        { sheetName: "Sheet3", range: "A1:A1", type: RangeType.FormulaRange, importance: 50, rowCount: 1, columnCount: 1 }
      ],
      rangeIdToChunkId: new Map()
    };

    const chunks = detector.createRangeChunks(sheet, detection as any);
    expect(chunks).toHaveLength(1);
    const chunk = chunks[0];
    expect(chunk.id).toBe(formatRangeId(detection.ranges[0] as any));
    expect(chunk.payload.values).toEqual([[10]]);
    expect(chunk.payload.formulas).toEqual([["=A1"]]);
    expect(chunk.refs).toContain("Sheet:Sheet3");
  });
});
