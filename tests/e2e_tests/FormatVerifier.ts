/**
 * Format Verifier
 * Provides utilities to verify Excel formatting properties
 */
// Office.js is available globally in the Excel add-in environment
declare const Excel: any;

/**
 * Expected cell formatting properties
 */
export interface ExpectedCellFormat {
  sheetName: string;
  cellAddress: string;
  properties: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
    fontColor?: string;
    fillColor?: string;
    numberFormat?: string;
    horizontalAlignment?: string; // Using string instead of Excel.HorizontalAlignment for compatibility
    verticalAlignment?: string; // Using string instead of Excel.VerticalAlignment for compatibility
    borders?: {
      top?: { style?: string; color?: string; weight?: string };
      bottom?: { style?: string; color?: string; weight?: string };
      left?: { style?: string; color?: string; weight?: string };
      right?: { style?: string; color?: string; weight?: string };
    };
    wrapText?: boolean;
    indentLevel?: number;
  };
}

/**
 * Expected chart formatting properties
 */
export interface ExpectedChartFormat {
  sheetName: string;
  chartName?: string;
  chartIndex?: number; // Use index if name is not available
  properties: {
    chartType?: string; // Using string instead of Excel.ChartType for compatibility
    title?: {
      text?: string;
      visible?: boolean;
      fontSize?: number;
      bold?: boolean;
    };
    legend?: {
      position?: string; // Using string instead of Excel.ChartLegendPosition for compatibility
      visible?: boolean;
    };
    hasDataLabels?: boolean;
    dataLabelsPosition?: string; // Using string instead of Excel.ChartDataLabelPosition for compatibility
    seriesCount?: number;
    height?: number;
    width?: number;
  };
}

/**
 * Format verification results
 */
export interface FormatVerificationResult {
  address: string;
  property: string;
  expected: any;
  actual: any;
  match: boolean;
}

/**
 * Format Verifier class
 * Provides methods to verify Excel formatting properties
 */
export class FormatVerifier {
  /**
   * Verify cell formatting properties
   * @param expectedFormats The expected cell formatting properties
   * @returns The verification results
   */
  public static async verifyCellFormatting(
    expectedFormats: ExpectedCellFormat[]
  ): Promise<FormatVerificationResult[]> {
    const results: FormatVerificationResult[] = [];

    await Excel.run(async (context) => {
      for (const format of expectedFormats) {
        try {
          const sheet = context.workbook.worksheets.getItem(format.sheetName);
          const range = sheet.getRange(format.cellAddress);
          
          range.format.fill.load("color");
          range.format.font.load("bold");
          range.format.font.load("italic");
          range.format.font.load("underline");
          range.format.font.load("size");
          range.format.font.load("color");
          range.format.load("horizontalAlignment");
          range.format.load("verticalAlignment");
          range.format.load("wrapText");
          range.format.load("indentLevel");
          range.load("numberFormat");
          
          // Load border properties if needed
          if (format.properties.borders) {
            range.format.borders.load("items/style");
            range.format.borders.load("items/color");
            range.format.borders.load("items/weight");
          }
          
          await context.sync();
          
          // Verify each property
          for (const [key, expectedValue] of Object.entries(format.properties)) {
            if (key === "borders" && expectedValue) {
              // Verify border properties
              for (const [borderSide, borderProps] of Object.entries(expectedValue)) {
                if (borderProps) {
                  const border = range.format.borders.getItem(borderSide as Excel.BorderIndex);
                  
                  for (const [borderProp, borderValue] of Object.entries(borderProps)) {
                    const actualValue = border[borderProp];
                    const match = this.compareValues(actualValue, borderValue);
                    
                    results.push({
                      address: `${format.sheetName}!${format.cellAddress}`,
                      property: `borders.${borderSide}.${borderProp}`,
                      expected: borderValue,
                      actual: actualValue,
                      match
                    });
                  }
                }
              }
            } else if (key === "fillColor") {
              // Handle fill color
              const actualValue = range.format.fill.color;
              const match = this.compareColors(actualValue, expectedValue as string);
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: "fillColor",
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "fontColor") {
              // Handle font color
              const actualValue = range.format.font.color;
              const match = this.compareColors(actualValue, expectedValue as string);
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: "fontColor",
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "bold" || key === "italic" || key === "underline") {
              // Handle font properties
              const actualValue = range.format.font[key];
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: key,
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "fontSize") {
              // Handle font size
              const actualValue = range.format.font.size;
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: key,
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "horizontalAlignment" || key === "verticalAlignment") {
              // Handle alignment
              const actualValue = range.format[key];
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: key,
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "numberFormat") {
              // Handle number format
              const actualValue = range.numberFormat;
              const match = this.compareNumberFormats(actualValue, expectedValue as string);
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: key,
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "wrapText" || key === "indentLevel") {
              // Handle other format properties
              const actualValue = range.format[key];
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!${format.cellAddress}`,
                property: key,
                expected: expectedValue,
                actual: actualValue,
                match
              });
            }
          }
        } catch (error) {
          console.error(`Error verifying format for ${format.sheetName}!${format.cellAddress}:`, error);
          
          // Add error result
          results.push({
            address: `${format.sheetName}!${format.cellAddress}`,
            property: "error",
            expected: "No error",
            actual: error.message,
            match: false
          });
        }
      }
    });
    
    return results;
  }
  
  /**
   * Verify chart formatting properties
   * @param expectedFormats The expected chart formatting properties
   * @returns The verification results
   */
  public static async verifyChartFormatting(
    expectedFormats: ExpectedChartFormat[]
  ): Promise<FormatVerificationResult[]> {
    const results: FormatVerificationResult[] = [];
    
    await Excel.run(async (context) => {
      for (const format of expectedFormats) {
        try {
          const sheet = context.workbook.worksheets.getItem(format.sheetName);
          let chart: Excel.Chart;
          
          // Handle the case where a string address is provided (e.g., "B3")
          if (format.chartName) {
            chart = sheet.charts.getItem(format.chartName);
          } else if (format.chartIndex !== undefined) {
            const charts = sheet.charts;
            charts.load("items");
            await context.sync();
            
            if (format.chartIndex >= 0 && format.chartIndex < charts.items.length) {
              chart = charts.items[format.chartIndex];
            } else {
              throw new Error(`Chart index ${format.chartIndex} out of range`);
            }
          } else {
            throw new Error("Either chartName or chartIndex must be provided");
          }
          
          // Load chart properties
          chart.load("chartType");
          chart.load("height");
          chart.load("width");
          chart.dataLabels.load("showValue"); // Using showValue instead of visible
          chart.dataLabels.load("position");
          chart.series.load("count");
          
          // Load title properties if needed
          if (format.properties.title) {
            chart.title.load("text");
            chart.title.load("visible");
            chart.title.format.font.load("bold");
            chart.title.format.font.load("size");
          }
          
          // Load legend properties if needed
          if (format.properties.legend) {
            chart.legend.load("position");
            chart.legend.load("visible");
          }
          
          await context.sync();
          
          // Verify each property
          for (const [key, expectedValue] of Object.entries(format.properties)) {
            if (key === "chartType") {
              const actualValue = chart.chartType;
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                property: "chartType",
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "title" && expectedValue) {
              // Verify title properties
              for (const [titleProp, titleValue] of Object.entries(expectedValue)) {
                let actualValue: any;
                let match: boolean;
                
                if (titleProp === "text") {
                  actualValue = chart.title.text;
                  match = actualValue === titleValue;
                } else if (titleProp === "visible") {
                  actualValue = chart.title.visible;
                  match = actualValue === titleValue;
                } else if (titleProp === "fontSize") {
                  actualValue = chart.title.format.font.size;
                  match = actualValue === titleValue;
                } else if (titleProp === "bold") {
                  actualValue = chart.title.format.font.bold;
                  match = actualValue === titleValue;
                } else {
                  continue;
                }
                
                results.push({
                  address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                  property: `title.${titleProp}`,
                  expected: titleValue,
                  actual: actualValue,
                  match
                });
              }
            } else if (key === "legend" && expectedValue) {
              // Verify legend properties
              for (const [legendProp, legendValue] of Object.entries(expectedValue)) {
                let actualValue: any;
                let match: boolean;
                
                if (legendProp === "position") {
                  actualValue = chart.legend.position;
                  match = actualValue === legendValue;
                } else if (legendProp === "visible") {
                  actualValue = chart.legend.visible;
                  match = actualValue === legendValue;
                } else {
                  continue;
                }
                
                results.push({
                  address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                  property: `legend.${legendProp}`,
                  expected: legendValue,
                  actual: actualValue,
                  match
                });
              }
            } else if (key === "hasDataLabels") {
              const actualValue = chart.dataLabels.showValue; // Using showValue instead of visible
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                property: "hasDataLabels",
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "dataLabelsPosition") {
              const actualValue = chart.dataLabels.position;
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                property: "dataLabelsPosition",
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "seriesCount") {
              const actualValue = chart.series.count;
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                property: "seriesCount",
                expected: expectedValue,
                actual: actualValue,
                match
              });
            } else if (key === "height" || key === "width") {
              const actualValue = chart[key];
              const match = actualValue === expectedValue;
              
              results.push({
                address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
                property: key,
                expected: expectedValue,
                actual: actualValue,
                match
              });
            }
          }
        } catch (error) {
          console.error(`Error verifying chart format in ${format.sheetName}:`, error);
          
          // Add error result
          results.push({
            address: `${format.sheetName}!Chart:${format.chartName || format.chartIndex}`,
            property: "error",
            expected: "No error",
            actual: error.message,
            match: false
          });
        }
      }
    });
    
    return results;
  }
  
  /**
   * Compare two values for equality
   * @param actual The actual value
   * @param expected The expected value
   * @returns True if the values match
   */
  private static compareValues(actual: any, expected: any): boolean {
    if (actual === expected) {
      return true;
    }
    
    // Handle special cases
    if (typeof expected === 'string' && typeof actual === 'string') {
      return actual.toLowerCase() === expected.toLowerCase();
    }
    
    return false;
  }
  
  /**
   * Compare two color values
   * @param actual The actual color
   * @param expected The expected color
   * @returns True if the colors match
   */
  private static compareColors(actual: string, expected: string): boolean {
    if (!actual || !expected) {
      return actual === expected;
    }
    
    // Normalize color values
    const normalizedActual = this.normalizeColor(actual);
    const normalizedExpected = this.normalizeColor(expected);
    
    return normalizedActual === normalizedExpected;
  }
  
  /**
   * Normalize a color value to a standard format
   * @param color The color value to normalize
   * @returns The normalized color value
   */
  private static normalizeColor(color: string): string {
    // Remove whitespace and convert to lowercase
    color = color.replace(/\s/g, '').toLowerCase();
    
    // Handle hex colors
    if (color.startsWith('#')) {
      // Expand shorthand hex (#rgb to #rrggbb)
      if (color.length === 4) {
        return `#${color[1]}${color[1]}${color[2]}${color[2]}${color[3]}${color[3]}`;
      }
      return color;
    }
    
    // Handle rgb/rgba colors
    if (color.startsWith('rgb')) {
      // Extract the RGB values
      const match = color.match(/(\d+),(\d+),(\d+)/);
      if (match) {
        const [_, r, g, b] = match;
        // Convert to hex
        return `#${Number(r).toString(16).padStart(2, '0')}${Number(g).toString(16).padStart(2, '0')}${Number(b).toString(16).padStart(2, '0')}`;
      }
    }
    
    // Return as is for named colors
    return color;
  }
  
  /**
   * Compare number formats
   * @param actual The actual number format
   * @param expected The expected number format
   * @returns True if the formats match
   */
  private static compareNumberFormats(actual: string, expected: string): boolean {
    if (actual === expected) {
      return true;
    }
    
    // Normalize number formats for comparison
    const normalizedActual = this.normalizeNumberFormat(actual);
    const normalizedExpected = this.normalizeNumberFormat(expected);
    
    return normalizedActual === normalizedExpected;
  }
  
  /**
   * Normalize a number format for comparison
   * @param format The number format to normalize
   * @returns The normalized format
   */
  private static normalizeNumberFormat(format: string): string {
    // Remove whitespace and convert to lowercase
    format = format.replace(/\s/g, '').toLowerCase();
    
    // Map common format strings to a canonical form
    const formatMap: Record<string, string> = {
      'general': 'general',
      'number': '0.00',
      'currency': '$#,##0.00',
      'accounting': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
      'shortdate': 'mm/dd/yyyy',
      'longdate': 'dddd, mmmm dd, yyyy',
      'time': 'hh:mm:ss am/pm',
      'percentage': '0.00%',
      'fraction': '# ?/?',
      'scientific': '0.00e+00',
      'text': '@'
    };
    
    // Return the canonical form if available
    return formatMap[format] || format;
  }
}
