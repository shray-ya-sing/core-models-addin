/**
 * Chart Models
 * Comprehensive models for Excel chart metadata and state tracking
 */

/**
 * Chart type enumeration
 * Represents all supported Excel chart types
 */
export enum ChartType {
  Area = 'area',
  AreaStacked = 'areaStacked',
  AreaStacked100 = 'areaStacked100',
  Bar = 'bar',
  BarStacked = 'barStacked',
  BarStacked100 = 'barStacked100',
  Column = 'column',
  ColumnStacked = 'columnStacked',
  ColumnStacked100 = 'columnStacked100',
  Line = 'line',
  LineMarkers = 'lineMarkers',
  LineStacked = 'lineStacked',
  LineStacked100 = 'lineStacked100',
  LineMarkersStacked = 'lineMarkersStacked',
  LineMarkersStacked100 = 'lineMarkersStacked100',
  Pie = 'pie',
  Doughnut = 'doughnut',
  Scatter = 'scatter',
  ScatterSmooth = 'scatterSmooth',
  ScatterSmoothNoMarkers = 'scatterSmoothNoMarkers',
  ScatterLines = 'scatterLines',
  ScatterLinesNoMarkers = 'scatterLinesNoMarkers',
  Radar = 'radar',
  RadarMarkers = 'radarMarkers',
  RadarFilled = 'radarFilled',
  StockHLC = 'stockHLC',
  StockOHLC = 'stockOHLC',
  StockVHLC = 'stockVHLC',
  StockVOHLC = 'stockVOHLC',
  Surface = 'surface',
  SurfaceWireframe = 'surfaceWireframe',
  SurfaceTopView = 'surfaceTopView',
  SurfaceTopViewWireframe = 'surfaceTopViewWireframe',
  Bubble = 'bubble',
  Bubble3D = 'bubble3D',
  Treemap = 'treemap',
  Sunburst = 'sunburst',
  Histogram = 'histogram',
  BoxWhisker = 'boxWhisker',
  Waterfall = 'waterfall',
  Funnel = 'funnel',
  Combo = 'combo'
}

/**
 * Chart legend position
 */
export enum ChartLegendPosition {
  Top = 'top',
  Bottom = 'bottom',
  Left = 'left',
  Right = 'right',
  Corner = 'corner',
  Custom = 'custom'
}

/**
 * Chart data label position
 */
export enum ChartDataLabelPosition {
  Center = 'center',
  InsideEnd = 'insideEnd',
  InsideBase = 'insideBase',
  OutsideEnd = 'outsideEnd',
  Above = 'above',
  Below = 'below',
  Left = 'left',
  Right = 'right',
  BestFit = 'bestFit',
  Custom = 'custom'
}

/**
 * Chart axis type
 */
export enum ChartAxisType {
  Category = 'category',
  Value = 'value',
  Series = 'series'
}

/**
 * Chart axis display unit
 */
export enum ChartAxisDisplayUnit {
  None = 'none',
  Hundreds = 'hundreds',
  Thousands = 'thousands',
  Millions = 'millions',
  Billions = 'billions',
  Trillions = 'trillions'
}

/**
 * Chart trendline type
 */
export enum ChartTrendlineType {
  Linear = 'linear',
  Exponential = 'exponential',
  Logarithmic = 'logarithmic',
  MovingAverage = 'movingAverage',
  Polynomial = 'polynomial',
  Power = 'power'
}

/**
 * Chart line style
 */
export enum ChartLineStyle {
  Continuous = 'continuous',
  Dash = 'dash',
  DashDot = 'dashDot',
  DashDotDot = 'dashDotDot',
  Dot = 'dot',
  LongDash = 'longDash',
  LongDashDot = 'longDashDot',
  LongDashDotDot = 'longDashDotDot',
  None = 'none'
}

/**
 * Chart marker style
 */
export enum ChartMarkerStyle {
  Automatic = 'automatic',
  Circle = 'circle',
  Dash = 'dash',
  Diamond = 'diamond',
  Dot = 'dot',
  None = 'none',
  Picture = 'picture',
  Plus = 'plus',
  Square = 'square',
  Star = 'star',
  Triangle = 'triangle',
  X = 'x'
}

/**
 * Chart formatting properties
 */
export interface ChartFormatting {
  fontColor?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fillColor?: string;
  lineColor?: string;
  lineStyle?: ChartLineStyle;
  lineWeight?: number;
}

/**
 * Chart title properties
 */
export interface ChartTitle {
  text: string;
  visible: boolean;
  overlay?: boolean;
  position?: string;
  format?: ChartFormatting;
}

/**
 * Chart legend properties
 */
export interface ChartLegend {
  visible: boolean;
  position: ChartLegendPosition;
  overlay?: boolean;
  format?: ChartFormatting;
}

/**
 * Chart data label properties
 */
export interface ChartDataLabel {
  visible: boolean;
  position?: ChartDataLabelPosition;
  showValue?: boolean;
  showSeriesName?: boolean;
  showCategoryName?: boolean;
  showLegendKey?: boolean;
  showPercentage?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  format?: ChartFormatting;
}

/**
 * Chart axis properties
 */
export interface ChartAxis {
  visible: boolean;
  title?: {
    text: string;
    visible: boolean;
    format?: ChartFormatting;
  };
  majorGridlines?: {
    visible: boolean;
    format?: {
      lineColor?: string;
      lineStyle?: ChartLineStyle;
      lineWeight?: number;
    };
  };
  minorGridlines?: {
    visible: boolean;
    format?: {
      lineColor?: string;
      lineStyle?: ChartLineStyle;
      lineWeight?: number;
    };
  };
  majorTickMarks?: string;
  minorTickMarks?: string;
  maximum?: number;
  minimum?: number;
  majorUnit?: number;
  minorUnit?: number;
  displayUnit?: ChartAxisDisplayUnit;
  showDisplayUnitLabel?: boolean;
  categoryType?: string;
  logScale?: boolean;
  reversePlotOrder?: boolean;
  scaleType?: string;
  position?: string;
  format?: ChartFormatting;
}

/**
 * Chart trendline properties
 */
export interface ChartTrendline {
  type: ChartTrendlineType;
  name?: string;
  polynomial?: number;
  period?: number;
  forward?: number;
  backward?: number;
  intercept?: number;
  displayEquation?: boolean;
  displayRSquared?: boolean;
  format?: {
    lineColor?: string;
    lineStyle?: ChartLineStyle;
    lineWeight?: number;
  };
}

/**
 * Chart series marker properties
 */
export interface ChartSeriesMarker {
  visible: boolean;
  style?: ChartMarkerStyle;
  size?: number;
  format?: {
    fillColor?: string;
    lineColor?: string;
    lineWeight?: number;
  };
}

/**
 * Chart series properties
 */
export interface ChartSeries {
  name: string;
  index: number;
  values: string; // Range address
  xValues?: string; // Range address
  plotOrder?: number;
  trendlines?: ChartTrendline[];
  dataLabels?: ChartDataLabel;
  markers?: ChartSeriesMarker;
  format?: {
    fillColor?: string;
    lineColor?: string;
    lineStyle?: ChartLineStyle;
    lineWeight?: number;
    transparency?: number;
  };
}

/**
 * Chart area properties
 */
export interface ChartArea {
  visible: boolean;
  format?: {
    fillColor?: string;
    lineColor?: string;
    lineStyle?: ChartLineStyle;
    lineWeight?: number;
  };
}

/**
 * Chart plot area properties
 */
export interface ChartPlotArea {
  visible: boolean;
  format?: {
    fillColor?: string;
    lineColor?: string;
    lineStyle?: ChartLineStyle;
    lineWeight?: number;
  };
}

/**
 * Chart state interface
 * Comprehensive representation of an Excel chart
 */
export interface ChartState {
  // Basic chart identification
  id: string;
  name: string;
  title: string;
  type: ChartType;
  
  // Chart location
  sheetName: string;
  topLeftCell: string;
  bottomRightCell?: string;
  
  // Chart dimensions
  height: number;
  width: number;
  
  // Chart source data
  sourceRange: string;
  
  // Chart elements
  chartTitle?: ChartTitle;
  legend?: ChartLegend;
  axes?: {
    categoryAxis?: ChartAxis;
    valueAxis?: ChartAxis;
    seriesAxis?: ChartAxis;
  };
  
  // Chart series
  series: ChartSeries[];
  
  // Chart areas
  chartArea?: ChartArea;
  plotArea?: ChartPlotArea;
  
  // Chart formatting
  hasDataTable?: boolean;
  dataTableFormat?: {
    showHorizontalBorder?: boolean;
    showVerticalBorder?: boolean;
    showOutlineBorder?: boolean;
    showLegendKeys?: boolean;
  };
  
  // Chart state
  is3D?: boolean;
  rotation?: {
    x?: number;
    y?: number;
    z?: number;
  };
  
  // Metadata
  createdAt: Date;
  lastModified: Date;
  lastCommand?: string;
}

/**
 * Chart reference for command context
 * Lightweight reference to a chart for use in command processing
 */
export interface ChartReference {
  id: string;
  name: string;
  title: string;
  sheetName: string;
  type: ChartType;
  isActive: boolean;
  lastAccessed: Date;
}

/**
 * Chart metadata chunk
 * For storing chart information in the metadata cache
 */
export interface ChartMetadataChunk {
  id: string;
  type: 'chart';
  etag: string;
  payload: ChartState;
  summary: string;
  refs: string[];
  lastCaptured: Date;
}
