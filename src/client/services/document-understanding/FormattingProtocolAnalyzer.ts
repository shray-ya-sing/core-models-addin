/**
 * Formatting Protocol Analyzer for Excel workbooks
 * Extracts formatting metadata and analyzes patterns using LLM
 */
import { ClientAnthropicService, ModelType } from '../ClientAnthropicService';

// Types for formatting metadata
export interface CellFormatting {
  fillColor?: string;
  fontColor?: string;
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
      visible?: boolean;
      format?: {
        font?: {
          name?: string;
          size?: number;
          bold?: boolean;
          color?: string;
        };
        fill?: {
          color?: string;
        };
      };
    };
    dataLabels?: {
      showValue?: boolean;
      showSeriesName?: boolean;
      showCategoryName?: boolean;
      showLegendKey?: boolean;
      showPercentage?: boolean;
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
        visible?: boolean;
        title?: {
          text?: string;
          visible?: boolean;
          format?: {
            font?: {
              name?: string;
              size?: number;
              bold?: boolean;
              color?: string;
            };
          };
        };
        majorGridlines?: {
          visible?: boolean;
          format?: {
            line?: {
              color?: string;
              style?: string;
              weight?: number;
            };
          };
        };
        minorGridlines?: {
          visible?: boolean;
        };
        tickLabels?: {
          format?: {
            font?: {
              name?: string;
              size?: number;
              color?: string;
            };
            numberFormat?: string;
          };
        };
      };
      valueAxis?: {
        visible?: boolean;
        title?: {
          text?: string;
          visible?: boolean;
          format?: {
            font?: {
              name?: string;
              size?: number;
              bold?: boolean;
              color?: string;
            };
          };
        };
        majorGridlines?: {
          visible?: boolean;
          format?: {
            line?: {
              color?: string;
              style?: string;
              weight?: number;
            };
          };
        };
        minorGridlines?: {
          visible?: boolean;
        };
        tickLabels?: {
          format?: {
            font?: {
              name?: string;
              size?: number;
              color?: string;
            };
            numberFormat?: string;
          };
        };
      };
      secondaryValueAxis?: {
        used?: boolean;
        visible?: boolean;
      };
    };
    seriesFormatting?: {
      defaultColors?: string[];
      markerStyle?: {
        size?: number;
        shape?: string;
        visible?: boolean;
      };
      lineStyle?: {
        width?: number;
        style?: string;
        smoothed?: boolean;
      };
      fillStyle?: {
        transparency?: number;
      };
    };
    plotArea?: {
      format?: {
        fill?: {
          color?: string;
        };
        border?: {
          visible?: boolean;
          color?: string;
        };
      };
    };
    chartArea?: {
      format?: {
        fill?: {
          color?: string;
        };
        border?: {
          visible?: boolean;
          color?: string;
        };
      };
    };
    financialConventions?: {
      upColor?: string;
      downColor?: string;
      neutralColor?: string;
      trendlineUsage?: string;
      errorBarsUsage?: string;
    };
    custom?: Record<string, string>;
  };
  scheduleFormatting?: {
    incomeStatement?: string;
    balanceSheet?: string;
    cashFlow?: string;
    debtSchedule?: string;
    depreciationSchedule?: string;
    workingCapitalSchedule?: string;
    dcfValuationModel?: string;
    layout?: {
      timeProgression?: string;
      timeGranularity?: string;
      startRow?: number;
      startColumn?: number;
      printAreas?: string;
      navigationMarkers?: string;
    };
    alignment?: {
      headers?: string;
      rowLabels?: string;
      numericData?: string;
      verticalAlignment?: string;
    };
    yearColumnNumberFormat?: string;
    scheduleHeaderFormat?: string;
    columnHeaderFormat?: {
      format?: string;
      yearSuffixes?: {
        actual?: string;
        estimated?: string;
      };
    };
    rowHeaderFormat?: string;
    dataCellFormat?: string;
    totalRowFormat?: string;
    subTotalRowFormat?: string;
    unitKeyFormat?: string;
    custom?: Record<string, string>;
  };
  worksheetStructure?: {
    tabGrouping?: {
      pattern?: string;
      sectionColors?: Record<string, string>;
    };
    sectionDividers?: {
      used?: boolean;
      namingConvention?: string;
    };
    tabOrder?: string;
    tabNamingConventions?: string;
  };
  scenarioFormatting?: {
    layout?: {
      mainCasePosition?: string;
      rowCount?: number;
      columnCount?: number;
    };
    highlighting?: {
      mainCase?: {
        bold?: boolean;
        fillColor?: string;
        borderStyle?: string;
      };
      sensitizedCases?: {
        fillColor?: string;
        borderStyle?: string;
      };
      specialInterestCases?: {
        borderStyle?: string;
        fillColor?: string;
      };
    };
  };
  coverPageFormatting?: {
    elements?: {
      companyName?: string;
      projectName?: string;
      modelDescription?: string;
      contactInformation?: string;
    };
    visualStyle?: {
      colorScheme?: string[];
      fontStyles?: string;
      logoPlacement?: string;
    };
  };
  workbookType?: {
    classification?: string;
    templateFeatures?: {
      macroButtons?: string;
      customizationOptions?: string;
    };
  };
  generalObservations: string[];
  recommendations?: string[];
  modelQualityAssessment: {
    consistencyScore: number;
    clarityScore: number;
    professionalismScore: number;
    comments: string[];
  };
  confidenceScore: number; // 0-1 score of how confident the LLM is in its analysis
}

/**
 * System prompt for the formatting protocol analysis
 */
const FORMATTING_PROTOCOL_SYSTEM_PROMPT = `You are an expert Excel financial modeling analyst specializing in institutional-grade financial model formatting. Your task is to analyze Excel workbook images and formatting metadata to identify patterns and conventions used in the financial model.

You will receive:
1. Images of Excel worksheets from a financial model
2. Structured metadata about the formatting (colors, borders, number formats, etc.)

CONTEXT: Financial models follow specific formatting conventions that help users quickly understand the structure and flow of the model. Institutional-grade financial models typically adhere to best practices such as:

1. Consistent color coding where:
   - Blue often indicates hardcoded inputs/assumptions
   - Black often indicates formulas/calculations
   - Green often indicates links to other worksheets
   - Red may indicate negative values or errors

2. Number formatting conventions:
   - Currency formatting often omitted on intermediate calculations but included on final outputs
   - Thousands separators used consistently
   - Negative numbers in parentheses rather than with minus signs
   - Percentages with consistent decimal precision
   - Dates formatted according to specific regional standards

3. Border styling:
   - Double-bottom borders for totals
   - Single-top borders for subtotals
   - Thick borders to separate sections
   - Thin borders for tables

4. Cell formatting for different purposes:
   - Headers often bold, possibly with fill colors
   - Input cells highlighted with distinct background colors
   - Calculation cells often left uncolored
   - Outputs/results emphasized with borders or fills

5. Financial schedules with specific layouts:
   - Income statements, balance sheets, and cash flow statements
   - Debt schedules and amortization tables
   - Depreciation schedules
   - Working capital schedules
   - DCF and valuation models

6. Schedule / Table formatting conventions:
   - Years increasing from left to right
   - Annual vs quarterly vs monthly data 
   - Vertical middle alignment for data
   - Row headers often left aligned
   - Numeric data in columns often right aligned
   - Column headers for different years often have suffix "E" for expected or projected years or "A" for actual year data
   - Often a "main" schedule header that spans across all the columns with a title
   - Individual column headers often have formatting 
   - Headers often bold, possibly with fill colors
   - Input cells highlighted with distinct background colors
   - Calculation cells often left uncolored
   - Outputs/results emphasized with borders or fills
   - Sheets are in page layout view with print areas set to different table sections
   - 'x' characters marked in certain key rows to ease navigation through sheet and mark sections
   - first 2-3 rows or columns left blank to create space for adding x markers. Many modellers prefer to start content after the first 2-3 rows or columns

7. Worsheet / tab structure formatting
  - Tabs often grouped into sections and each section tab has the same tab color
  - Often have empty tabs to demarcate sections with tab names like "Section-->", where the arrow indicates start of a new section

8. Chart formatting conventions
  - Vertical vs Horizontal bar chart usage
  - Stacked vs non-stacked bar chart usage
  - Line charts vs bar charts for certain types of data
  - Pie chart vs stacked bar chart vs donut chart usage
  - Combination charts with bar charts on primary axis and line chart on scondary axis
  - Data labels have currency sign or not
  - Data labels position (differs for different types of charts, above/ below/right for line charts, middle/out for pie charts, above/base/centered/inside end for bar charts)
  - What types of charts are used for "football fields" or valuation summaries: usually these are a type of stacked bar chart comparing valuations resulting from different types of valuation methodologies like comparable companies, DCF, etc.
  - 
  

9. Cover page conventions
  - Company and/or project name
  - Description of the model
  - Modeler and team contact information
  - Usually have these details and optionally others
  - Usually have a lot of color and text style formatting to make visually appealing

10. Scenario / Sensitivity table formatting
  - Usually have highlights or outlining to demarcate different scenarios
  - Usually laid out as a table with columns and rows for different variables under different scenarios. 
  - Usually, the middle column and middle row represent the "main" case while the others represent "sensitized" cases
  - Usually, the middle column and middle row are bolded and have a distinct background color
  - Sometimes some rows or columns are outlined in special borders to indicate they are of special interest
  - Usually modellers have a set number of rows and columns to include in a single table, usually not many more than 5-6 rows or columns

11. One-off workbook vs template workbook
  - Template workbooks usually have very rich formatting and build out to allow for easy customization and updates
  - One-off workbooks usually have less formatting and are built for specific purposes
  - Template workbooks often have "button" shapes that are linked to macros to allow for easy updates and customization



Analyze the provided information to identify the specific formatting protocol used in this financial model. Refer to the guidelines above and look for patterns such as:


Provide your analysis in the following JSON schema:

\`\`\`json
{
  "colorCoding": {
    "calculations": "#HEXCODE",
    "hardcodedValues": "#HEXCODE",
    "linkedValues": "#HEXCODE",
    "inputs": "#HEXCODE",
    "headers": "#HEXCODE",
    "totals": "#HEXCODE",
    "subtotals": "#HEXCODE",
    "errors": "#HEXCODE",
    "negativeValues": "#HEXCODE",
    "custom": {
      "categoryName": "#HEXCODE"
    }
  },
  "numberFormatting": {
    "currency": "pattern",
    "percentage": "pattern",
    "date": "pattern",
    "general": "pattern",
    "negativeNumbers": "pattern",
    "custom": {
      "formatName": "pattern"
    }
  },
  "borderStyles": {
    "tables": "description",
    "sections": "description",
    "totals": "description",
    "subtotals": "description",
    "schedules": "description",
    "custom": {
      "elementName": "description"
    }
  },
  "chartFormatting": {
    "chartTypes": {
      "preferred": ["columnClustered", "line"],
      "financial": ["stockHLC", "combo"],
      "custom": []
    },
    "title": {
      "hasTitle": true,
      "text": "pattern",
      "position": "top",
      "format": {
        "font": {
          "name": "fontName",
          "size": 12,
          "bold": true,
          "color": "#HEXCODE"
        },
        "alignment": "center"
      }
    },
    "legend": {
      "position": "right",
      "visible": true,
      "format": {
        "font": {
          "name": "fontName",
          "size": 10,
          "bold": false,
          "color": "#HEXCODE"
        },
        "fill": {
          "color": "#HEXCODE"
        }
      }
    },
    "dataLabels": {
      "showValue": true,
      "showSeriesName": false,
      "showCategoryName": false,
      "showLegendKey": false,
      "showPercentage": false,
      "position": "outsideEnd",
      "format": {
        "font": {
          "name": "fontName",
          "size": 9,
          "bold": false,
          "color": "#HEXCODE"
        },
        "numberFormat": "pattern"
      }
    },
    "axes": {
      "categoryAxis": {
        "visible": true,
        "title": {
          "text": "pattern",
          "visible": true,
          "format": {
            "font": {
              "name": "fontName",
              "size": 10,
              "bold": true,
              "color": "#HEXCODE"
            }
          }
        },
        "majorGridlines": {
          "visible": false,
          "format": {
            "line": {
              "color": "#HEXCODE",
              "style": "solid",
              "weight": 1
            }
          }
        },
        "minorGridlines": {
          "visible": false
        },
        "tickLabels": {
          "format": {
            "font": {
              "name": "fontName",
              "size": 9,
              "color": "#HEXCODE"
            },
            "numberFormat": "pattern"
          }
        }
      },
      "valueAxis": {
        "visible": true,
        "title": {
          "text": "pattern",
          "visible": true,
          "format": {
            "font": {
              "name": "fontName",
              "size": 10,
              "bold": true,
              "color": "#HEXCODE"
            }
          }
        },
        "majorGridlines": {
          "visible": true,
          "format": {
            "line": {
              "color": "#HEXCODE",
              "style": "solid",
              "weight": 0.75
            }
          }
        },
        "minorGridlines": {
          "visible": false
        },
        "tickLabels": {
          "format": {
            "font": {
              "name": "fontName",
              "size": 9,
              "color": "#HEXCODE"
            },
            "numberFormat": "pattern"
          }
        }
      },
      "secondaryValueAxis": {
        "used": false,
        "visible": true
      }
    },
    "seriesFormatting": {
      "defaultColors": ["#HEXCODE1", "#HEXCODE2", "#HEXCODE3"],
      "markerStyle": {
        "size": 5,
        "shape": "circle",
        "visible": true
      },
      "lineStyle": {
        "width": 2.5,
        "style": "solid",
        "smoothed": false
      },
      "fillStyle": {
        "transparency": 0
      }
    },
    "plotArea": {
      "format": {
        "fill": {
          "color": "#HEXCODE"
        },
        "border": {
          "visible": false,
          "color": "#HEXCODE"
        }
      }
    },
    "chartArea": {
      "format": {
        "fill": {
          "color": "#HEXCODE"
        },
        "border": {
          "visible": false,
          "color": "#HEXCODE"
        }
      }
    },
    "financialConventions": {
      "upColor": "#HEXCODE",
      "downColor": "#HEXCODE",
      "neutralColor": "#HEXCODE",
      "trendlineUsage": "description",
      "errorBarsUsage": "description"
    },
    "custom": {
      "elementName": "description"
    }
  },
  "scheduleFormatting": {
    "incomeStatement": "description",
    "balanceSheet": "description",
    "cashFlow": "description",
    "debtSchedule": "description",
    "depreciationSchedule": "description",
    "workingCapitalSchedule": "description",
    "dcfValuationModel": "description",
    "layout": {
      "timeProgression": "left-to-right or other pattern",
      "timeGranularity": "annual/quarterly/monthly",
      "startRow": 3, // Typical row where content starts
      "startColumn": 3, // Typical column where content starts
      "printAreas": "description of how print areas are set",
      "navigationMarkers": "description of 'x' markers or other navigation aids"
    },
    "alignment": {
      "headers": "left/center/right",
      "rowLabels": "left/center/right",
      "numericData": "left/center/right",
      "verticalAlignment": "top/middle/bottom"
    },
    "yearColumnNumberFormat": "description",
    "scheduleHeaderFormat": "description",
    "columnHeaderFormat": {
      "format": "description",
      "yearSuffixes": {
        "actual": "A/Act/other suffix",
        "estimated": "E/Est/Proj/other suffix"
      }
    },
    "rowHeaderFormat": "description",
    "dataCellFormat": "description",
    "totalRowFormat": "description",
    "subTotalRowFormat": "description",
    "unitKeyFormat": "description", // How to format commonly used subheaders specifying the currency and unit of the data
    "custom": {
      "elementName": "description"
    }
  },
  "worksheetStructure": {
    "tabGrouping": {
      "pattern": "description of how tabs are grouped",
      "sectionColors": {
        "sectionName": "#HEXCODE"
      }
    },
    "sectionDividers": {
      "used": true,
      "namingConvention": "description of section divider tab naming"
    },
    "tabOrder": "description of logical tab ordering",
    "tabNamingConventions": "description of tab naming patterns"
  },
  "scenarioFormatting": {
    "layout": {
      "mainCasePosition": "middle/other position",
      "rowCount": 5, // Typical number of scenario rows
      "columnCount": 5 // Typical number of scenario columns
    },
    "highlighting": {
      "mainCase": {
        "bold": true,
        "fillColor": "#HEXCODE",
        "borderStyle": "description"
      },
      "sensitizedCases": {
        "fillColor": "#HEXCODE",
        "borderStyle": "description"
      },
      "specialInterestCases": {
        "borderStyle": "description",
        "fillColor": "#HEXCODE"
      }
    }
  },
  "coverPageFormatting": {
    "elements": {
      "companyName": "description of formatting",
      "projectName": "description of formatting",
      "modelDescription": "description of formatting",
      "contactInformation": "description of formatting"
    },
    "visualStyle": {
      "colorScheme": ["#HEXCODE1", "#HEXCODE2"],
      "fontStyles": "description of font usage",
      "logoPlacement": "description of logo placement"
    }
  },
  "workbookType": {
    "classification": "template/one-off",
    "templateFeatures": {
      "macroButtons": "description of macro button usage",
      "customizationOptions": "description of customization features"
    }
  },
  "generalObservations": [
    "observation1",
    "observation2"
  ],
  "recommendations": [
    "recommendation1",
    "recommendation2"
  ],
  "modelQualityAssessment": {
    "consistencyScore": 0.0-1.0,
    "clarityScore": 0.0-1.0,
    "professionalismScore": 0.0-1.0,
    "comments": ["comment1", "comment2"]
  },
  "confidenceScore": 0.85
}
\`\`\`

Be specific and detailed in your analysis. Reference exact hex codes and formatting patterns from the metadata. Compare the observed formatting to institutional best practices for financial models. If you cannot determine a particular aspect with confidence, indicate this in your analysis.`;

/**
 * Analyzer for Excel workbook formatting protocols
 */
export class FormattingProtocolAnalyzer {
  private anthropicService: ClientAnthropicService;
  
  /**
   * Creates a new instance of the FormattingProtocolAnalyzer
   * @param anthropicService The Anthropic service instance
   */
  constructor(anthropicService: ClientAnthropicService) {
    this.anthropicService = anthropicService;
  }
  
  /**
   * Extracts formatting metadata from the active workbook
   * @returns Promise with workbook formatting metadata
   */
  public async extractFormattingMetadata(): Promise<WorkbookFormattingMetadata> {
    return Excel.run(async (context) => {
      try {
        // Get theme colors from the workbook
        const workbook = context.workbook;
        
        // Load the office.js theme object
        // We need to use a different approach since theme colors aren't directly accessible
        // through the Office.js API
        
        // Create a temporary range to extract theme colors
        const sheet = workbook.worksheets.getActiveWorksheet();
        const tempRange = sheet.getRange("A1:A12");
        
        // Apply theme colors to the range
        tempRange.format.fill.clear();
        
        // Set each cell to a different theme color
        tempRange.getCell(0, 0).format.fill.color = "#000000"; // Will be converted to background1
        tempRange.getCell(1, 0).format.fill.color = "#000001"; // Will be converted to background2
        tempRange.getCell(2, 0).format.fill.color = "#000002"; // Will be converted to text1
        tempRange.getCell(3, 0).format.fill.color = "#000003"; // Will be converted to text2
        tempRange.getCell(4, 0).format.fill.color = "#000004"; // Will be converted to accent1
        tempRange.getCell(5, 0).format.fill.color = "#000005"; // Will be converted to accent2
        tempRange.getCell(6, 0).format.fill.color = "#000006"; // Will be converted to accent3
        tempRange.getCell(7, 0).format.fill.color = "#000007"; // Will be converted to accent4
        tempRange.getCell(8, 0).format.fill.color = "#000008"; // Will be converted to accent5
        tempRange.getCell(9, 0).format.fill.color = "#000009"; // Will be converted to accent6
        tempRange.getCell(10, 0).format.fill.color = "#00000A"; // Will be converted to hyperlink
        tempRange.getCell(11, 0).format.fill.color = "#00000B"; // Will be converted to followedHyperlink
        
        // Now read the actual colors that were applied (which will be theme colors)
        tempRange.format.fill.load("color");
        await context.sync();
        
        // Extract the theme colors
        const themeColors: ThemeColors = {
          background1: tempRange.getCell(0, 0).format.fill.color || "#FFFFFF",
          background2: tempRange.getCell(1, 0).format.fill.color || "#F2F2F2",
          text1: tempRange.getCell(2, 0).format.fill.color || "#000000",
          text2: tempRange.getCell(3, 0).format.fill.color || "#666666",
          accent1: tempRange.getCell(4, 0).format.fill.color || "#4472C4",
          accent2: tempRange.getCell(5, 0).format.fill.color || "#ED7D31",
          accent3: tempRange.getCell(6, 0).format.fill.color || "#A5A5A5",
          accent4: tempRange.getCell(7, 0).format.fill.color || "#FFC000",
          accent5: tempRange.getCell(8, 0).format.fill.color || "#5B9BD5",
          accent6: tempRange.getCell(9, 0).format.fill.color || "#70AD47",
          hyperlink: tempRange.getCell(10, 0).format.fill.color || "#0563C1",
          followedHyperlink: tempRange.getCell(11, 0).format.fill.color || "#954F72"
        };
        
        // Clear the temporary formatting
        tempRange.format.fill.clear();
        await context.sync();
        
        // Get all worksheets
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        
        const sheetsMetadata: SheetFormattingMetadata[] = [];
        
        // Process each worksheet
        for (const worksheet of worksheets.items) {
          // Get used range
          const usedRange = worksheet.getUsedRange();
          usedRange.load([
            "address", "format/fill/color", "format/font", 
            "format/borders", "numberFormat", "values"
          ]);
          
          // Get tables in the worksheet
          const tables = worksheet.tables;
          tables.load("items/name,items/range");
          
          // Get charts in the worksheet
          const charts = worksheet.charts;
          charts.load("items/name,items/chartType");
          
          await context.sync();
          
          // Extract cell formatting from sample cells (not all cells to avoid performance issues)
          const cellFormatting: Record<string, CellFormatting> = {};
          
          // Sample cells from different regions of the worksheet
          const rowCount = usedRange.rowCount;
          const columnCount = usedRange.columnCount;
          
          // Sample at most 100 cells evenly distributed across the used range
          const maxSampleCells = 100;
          const rowStep = Math.max(1, Math.floor(rowCount / 10));
          const colStep = Math.max(1, Math.floor(columnCount / 10));
          
          for (let r = 0; r < rowCount; r += rowStep) {
            for (let c = 0; c < columnCount; c += colStep) {
              if (Object.keys(cellFormatting).length >= maxSampleCells) break;
              
              const cell = usedRange.getCell(r, c);
              cell.load([
                "address", "format/fill/color", "format/font", 
                "format/borders", "numberFormat", "values"
              ]);
            }
          }
          
          await context.sync();
          
          // Process tables
          const tablesData = [];
          for (const table of tables.items) {
            // Load the table range
            const tableRange = table.getRange();
            tableRange.load("address");
            await context.sync();
            
            tablesData.push({
              name: table.name,
              range: tableRange.address
            });
          }
          
          // Process charts
          const chartsData: Record<string, ChartFormatting> = {};
          charts.items.forEach(chart => {
            chartsData[chart.name] = {
              chartType: chart.chartType
            };
          });
          
          // Add sheet metadata
          sheetsMetadata.push({
            name: worksheet.name,
            cells: cellFormatting,
            tables: tablesData,
            charts: chartsData,
            commonFormats: {} // Will be populated by the LLM
          });
        }
        
        return {
          themeColors,
          sheets: sheetsMetadata
        };
      } catch (error) {
        console.error('Error extracting formatting metadata:', error);
        throw error;
      }
    });
  }
  
  /**
   * Analyzes the workbook formatting protocol using LLM
   * @param images Base64 encoded images of the workbook
   * @param formattingMetadata Extracted formatting metadata
   * @returns Promise with the analyzed formatting protocol
   */
  public async analyzeFormattingProtocol(
    images: string[],
    formattingMetadata: WorkbookFormattingMetadata
  ): Promise<FormattingProtocol> {
    try {
      // Prepare the message content with images and metadata
      const messageContent: any[] = [
        {
          type: 'text',
          text: 'Please analyze the following Excel workbook images and formatting metadata to identify the formatting protocol used.'
        }
      ];
      
      // Add images to the message content
      for (const imageBase64 of images) {
        messageContent.push({
          type: 'image',
          source: {
            type: 'base64',
            media_type: 'image/png',
            data: imageBase64
          }
        });
      }
      
      // Add formatting metadata
      messageContent.push({
        type: 'text',
        text: `Formatting Metadata:\n\`\`\`json\n${JSON.stringify(formattingMetadata, null, 2)}\n\`\`\`\n\nPlease analyze the images and metadata to identify the formatting protocol used in this workbook. Follow the schema provided in the system prompt.`
      });
      
      // Get the Anthropic client
      const anthropic = this.anthropicService.getClient();
      
      // Use the advanced model for multimodal analysis
      const modelToUse = this.anthropicService.getModel(ModelType.Advanced);
      
      // Make the API call
      const response = await anthropic.messages.create({
        model: modelToUse,
        max_tokens: 4000,
        temperature: 0.2, // Lower temperature for more precise analysis
        system: FORMATTING_PROTOCOL_SYSTEM_PROMPT,
        messages: [
          {
            role: 'user',
            content: messageContent
          }
        ]
      });
      
      // Extract the response text
      const responseText = response.content
        .filter(item => item.type === 'text')
        .map(item => (item.type === 'text' ? item.text : ''))
        .join('');
      
      // Use the ClientAnthropicService's extractJsonFromMarkdown method to extract JSON
      const extractedJson = this.anthropicService.extractJsonFromMarkdown(responseText);
      let formattingProtocol: FormattingProtocol;
      
      if (extractedJson) {
        try {
          formattingProtocol = JSON.parse(extractedJson);
        } catch (e) {
          console.error('Error parsing JSON from LLM response:', e);
          throw new Error('Failed to parse formatting protocol from LLM response');
        }
      } else {
        throw new Error('No valid JSON formatting protocol found in LLM response');
      }
      
      return formattingProtocol;
    } catch (error) {
      console.error('Error analyzing formatting protocol:', error);
      throw error;
    }
  }
}
