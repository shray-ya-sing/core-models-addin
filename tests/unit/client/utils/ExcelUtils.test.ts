/**
 * Unit tests for ExcelUtils
 */

import { 
  toA1, 
  columnToLetter, 
  letterToColumn, 
  parseA1, 
  parseReference, 
  createReference, 
  createRangeAddress 
} from "../../../../src/client/utils/ExcelUtils";

describe("ExcelUtils", () => {
  // Tests for toA1 function
  describe("toA1", () => {
    test("converts (0,0) to A1", () => {
      expect(toA1(0, 0)).toBe("A1");
    });
    
    test("converts (5,2) to C6", () => {
      expect(toA1(5, 2)).toBe("C6");
    });
    
    test("handles large column values", () => {
      expect(toA1(0, 26)).toBe("AA1");
      expect(toA1(0, 27)).toBe("AB1");
      expect(toA1(9, 51)).toBe("AZ10");
      expect(toA1(0, 701)).toBe("ZZ1");
    });
  });
  
  // Tests for columnToLetter function
  describe("columnToLetter", () => {
    test("converts 0 to A", () => {
      expect(columnToLetter(0)).toBe("A");
    });
    
    test("converts 25 to Z", () => {
      expect(columnToLetter(25)).toBe("Z");
    });
    
    test("converts 26 to AA", () => {
      expect(columnToLetter(26)).toBe("AA");
    });
    
    test("converts 27 to AB", () => {
      expect(columnToLetter(27)).toBe("AB");
    });
    
    test("converts 51 to AZ", () => {
      expect(columnToLetter(51)).toBe("AZ");
    });
    
    test("converts 701 to ZZ", () => {
      expect(columnToLetter(701)).toBe("ZZ");
    });
  });
  
  // Tests for letterToColumn function
  describe("letterToColumn", () => {
    test("converts A to 0", () => {
      expect(letterToColumn("A")).toBe(0);
    });
    
    test("converts Z to 25", () => {
      expect(letterToColumn("Z")).toBe(25);
    });
    
    test("converts AA to 26", () => {
      expect(letterToColumn("AA")).toBe(26);
    });
    
    test("converts AB to 27", () => {
      expect(letterToColumn("AB")).toBe(27);
    });
    
    test("converts AZ to 51", () => {
      expect(letterToColumn("AZ")).toBe(51);
    });
    
    test("converts ZZ to 701", () => {
      expect(letterToColumn("ZZ")).toBe(701);
    });
    
    test("works with lowercase letters", () => {
      expect(letterToColumn("aa")).toBe(26);
    });
  });
  
  // Tests for parseA1 function
  describe("parseA1", () => {
    test("parses A1 correctly", () => {
      const result = parseA1("A1");
      expect(result.row).toBe(0);
      expect(result.column).toBe(0);
    });
    
    test("parses C6 correctly", () => {
      const result = parseA1("C6");
      expect(result.row).toBe(5);
      expect(result.column).toBe(2);
    });
    
    test("parses AA1 correctly", () => {
      const result = parseA1("AA1");
      expect(result.row).toBe(0);
      expect(result.column).toBe(26);
    });
    
    test("parses ZZ100 correctly", () => {
      const result = parseA1("ZZ100");
      expect(result.row).toBe(99);
      expect(result.column).toBe(701);
    });
    
    test("throws error for invalid reference", () => {
      expect(() => parseA1("123")).toThrow();
      expect(() => parseA1("ABC")).toThrow();
      expect(() => parseA1("A1A")).toThrow();
    });
  });
  
  // Tests for parseReference function
  describe("parseReference", () => {
    test("parses Sheet1!A1 correctly", () => {
      const result = parseReference("Sheet1!A1");
      expect(result.sheet).toBe("Sheet1");
      expect(result.address).toBe("A1");
    });
    
    test("parses 'Sheet Name With Spaces'!B5 correctly", () => {
      const result = parseReference("'Sheet Name With Spaces'!B5");
      expect(result.sheet).toBe("'Sheet Name With Spaces'");
      expect(result.address).toBe("B5");
    });
    
    test("throws error for invalid reference", () => {
      expect(() => parseReference("Sheet1.A1")).toThrow();
      expect(() => parseReference("A1")).toThrow();
    });
  });
  
  // Tests for createReference function
  describe("createReference", () => {
    test("creates reference from sheet name and address", () => {
      expect(createReference("Sheet1", "A1")).toBe("Sheet1!A1");
    });
    
    test("works with sheet names containing spaces", () => {
      expect(createReference("'Sheet Name With Spaces'", "B5")).toBe("'Sheet Name With Spaces'!B5");
    });
  });
  
  // Tests for createRangeAddress function
  describe("createRangeAddress", () => {
    test("creates range address from top-left and bottom-right cells", () => {
      expect(createRangeAddress("A1", "B2")).toBe("A1:B2");
    });
    
    test("works with complex cell references", () => {
      expect(createRangeAddress("AA1", "ZZ100")).toBe("AA1:ZZ100");
    });
  });
  
  // Additional round-trip conversion tests
  describe("round-trip conversions", () => {
    test("columnToLetter and letterToColumn are inverses", () => {
      const letters = ["A", "Z", "AA", "ZZ", "BC"];
      letters.forEach(l => {
        expect(columnToLetter(letterToColumn(l))).toBe(l.toUpperCase());
      });
    });

    test("toA1 and parseA1 are inverses", () => {
      const coords = [
        { row: 0, col: 0 },
        { row: 5, col: 26 },
        { row: 99, col: 701 }
      ];
      coords.forEach(c => {
        const a1 = toA1(c.row, c.col);
        const parsed = parseA1(a1);
        expect(parsed.row).toBe(c.row);
        expect(parsed.column).toBe(c.col);
      });
    });

    test("createReference and parseReference are inverses", () => {
      const sheet = "Sheet Test";
      const address = "AB10";
      const ref = createReference(sheet, address);
      const parsed = parseReference(ref);
      expect(parsed.sheet).toBe(sheet);
      expect(parsed.address).toBe(address);
    });

    test("parses lowercase A1 correctly", () => {
      const result = parseA1("a1");
      expect(result.row).toBe(0);
      expect(result.column).toBe(0);
    });
  });
});
