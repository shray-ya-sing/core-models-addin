/**
 * Models for Excel workbook formatting analysis
 */

// Types for formatting metadata
export interface CellFormatting {
  fillColor?: string;
  fontColor?: string;
  fontName?: string;
  fontBold?: boolean;
  fontStyle?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    name?: string;
    size?: number;
  };
  borders?: {
    top?: { style?: string; color?: string };
    bottom?: { style?: string; color?: string };
    left?: { style?: string; color?: string };
    right?: { style?: string; color?: string };
  };
  numberFormat?: string;
  // Accept both string and array values to match Office.js API
  value?: string | any[] | any[][];
  alignment?: {
    horizontal?: string;
    vertical?: string;
  };
}

export interface ChartFormatting {
  chartType: string;
  title?: string;
  hasLegend?: boolean;
  legendPosition?: string;
  dataLabels?: boolean;
  seriesColors?: string[];
  axisFormatting?: {
    title?: string;
    gridlines?: boolean;
    numberFormat?: string;
  };
}

export interface ThemeColors {
  background1: string;
  background2: string;
  text1: string;
  text2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hyperlink: string;
  followedHyperlink: string;
}

export interface SheetFormattingMetadata {
  name: string;
  cells: Record<string, CellFormatting>;
  tables: Array<{
    name: string;
    range: string;
    headerRow?: boolean;
    totalRow?: boolean;
    headerFormatting?: CellFormatting;
    dataFormatting?: CellFormatting;
    totalRowFormatting?: CellFormatting;
  }>;
  charts: Record<string, ChartFormatting>;
  commonFormats: {
    headers?: CellFormatting;
    inputs?: CellFormatting;
    calculations?: CellFormatting;
    totals?: CellFormatting;
    subtotals?: CellFormatting;
  };
}

export interface WorkbookFormattingMetadata {
  themeColors: ThemeColors;
  sheets: SheetFormattingMetadata[];
}

export interface FormattingProtocol {
  colorCoding: {
    calculations?: string;
    hardcodedValues?: string;
    linkedValues?: string;
    inputs?: string;
    headers?: string;
    totals?: string;
    subtotals?: string;
    errors?: string;
    negativeValues?: string;
    custom?: Record<string, string>;
  };
  numberFormatting: {
    currency?: string;
    percentage?: string;
    date?: string;
    general?: string;
    negativeNumbers?: string;
    custom?: Record<string, string>;
  };
  borderStyles: {
    tables?: string;
    sections?: string;
    totals?: string;
    subtotals?: string;
    schedules?: string;
    custom?: Record<string, string>;
  };
  chartFormatting: {
    chartTypes?: {
      preferred?: string[];
      financial?: string[];
      custom?: string[];
    };
    title?: {
      hasTitle?: boolean;
      text?: string;
      position?: string;
      format?: {
        font?: {
          name?: string;
          size?: number;
          bold?: boolean;
          color?: string;
        };
        alignment?: string;
      };
    };
    legend?: {
      position?: string;
      format?: {
        font?: {
          name?: string;
          size?: number;
          bold?: boolean;
          color?: string;
        };
      };
    };
    dataLabels?: {
      visible?: boolean;
      position?: string;
      format?: {
        font?: {
          name?: string;
          size?: number;
          bold?: boolean;
          color?: string;
        };
        numberFormat?: string;
      };
    };
    axes?: {
      categoryAxis?: {
        title?: {
          visible?: boolean;
          text?: string;
          format?: {
            font?: {
              name?: string;
              size?: number;
              bold?: boolean;
              color?: string;
            };
          };
        };
        gridlines?: boolean;
        format?: {
          font?: {
            name?: string;
            size?: number;
            bold?: boolean;
            color?: string;
          };
        };
      };
      valueAxis?: {
        title?: {
          visible?: boolean;
          text?: string;
          format?: {
            font?: {
              name?: string;
              size?: number;
              bold?: boolean;
              color?: string;
            };
          };
        };
        gridlines?: boolean;
        format?: {
          font?: {
            name?: string;
            size?: number;
            bold?: boolean;
            color?: string;
          };
          numberFormat?: string;
        };
      };
    };
    seriesFormatting?: {
      colors?: string[];
      markers?: {
        visible?: boolean;
        style?: string;
      };
      lineStyles?: string[];
    };
  };
  fontUsage: {
    headers?: {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    };
    body?: {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    };
    titles?: {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    };
    custom?: Record<string, {
      name?: string;
      size?: number;
      bold?: boolean;
      color?: string;
    }>;
  };
  tableFormatting: {
    headerRow?: {
      fillColor?: string;
      fontColor?: string;
      fontBold?: boolean;
      borders?: string;
    };
    dataRows?: {
      alternatingColors?: boolean;
      evenRowColor?: string;
      oddRowColor?: string;
    };
    totalRow?: {
      fillColor?: string;
      fontColor?: string;
      fontBold?: boolean;
      borders?: string;
    };
  };
  scheduleFormatting: {
    incomeStatement?: string;
    balanceSheet?: string;
    cashFlow?: string;
    debtSchedule?: string;
    capex?: string;
    depreciation?: string;
    taxSchedule?: string;
    workingCapital?: string;
    custom?: Record<string, string>;
  };
  workbookStructure: {
    tabOrdering?: string;
    tabGrouping?: string;
    sectionDividers?: string;
    inputSections?: string;
    outputSections?: string;
    calculationSections?: string;
  };
  scenarioFormatting: {
    sensitivityTables?: {
      layout?: string;
      highlighting?: string;
      baseCase?: {
        position?: string;
        formatting?: string;
      };
    };
    scenarioManager?: {
      used?: boolean;
      structure?: string;
    };
    dataTables?: {
      used?: boolean;
      structure?: string;
    };
  };
  coverPageFormatting?: {
    title?: {
      font?: {
        name?: string;
        size?: number;
        bold?: boolean;
        color?: string;
      };
      alignment?: string;
    };
    subtitle?: {
      font?: {
        name?: string;
        size?: number;
        bold?: boolean;
        color?: string;
      };
      alignment?: string;
    };
    logo?: {
      position?: string;
      size?: string;
    };
    companyInfo?: {
      font?: {
        name?: string;
        size?: number;
        bold?: boolean;
        color?: string;
      };
      alignment?: string;
    };
    date?: {
      format?: string;
      position?: string;
    };
  };
  workbookType?: {
    financialModel?: boolean;
    threeStatementModel?: boolean;
    dcfModel?: boolean;
    lboModel?: boolean;
    mergersModel?: boolean;
    budgetModel?: boolean;
    forecastModel?: boolean;
    operationalModel?: boolean;
    dashboardModel?: boolean;
    custom?: string;
  };
  yearSuffixes?: {
    actual?: string;
    projected?: string;
    budget?: string;
  };
}

// Original state of cell fill for restoration
export interface OriginalFill {
  color: string | null;
  hasFill: boolean;
}
