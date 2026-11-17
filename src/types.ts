// Shared shapes of the presentation document, used by the main part, the panel and the runner.

export type ColorRole =
  | "background"
  | "surface"
  | "text"
  | "muted"
  | "line"
  | "accent"
  | "accent2"
  | "onAccent"
  | "chart1"
  | "chart2"
  | "chart3"
  | "chart4"
  | "chart5";
export type FontRole = "heading" | "body";
export type SizeRole = "display" | "h1" | "h2" | "h3" | "body" | "small" | "caption";
export type RoleField = "fill" | "stroke" | "color" | "font" | "size";
export type ObjectRoles = Partial<Record<RoleField, string>>;

export interface Theme {
  id: string | null;
  background: string;
  colors: Record<ColorRole, string>;
  fonts: Record<FontRole, string>;
  headingBold: boolean;
}

/** Old documents may store only a background. */
export interface StoredTheme {
  id?: string | null;
  background: string;
  colors?: Record<ColorRole, string>;
  fonts?: Record<FontRole, string>;
  headingBold?: boolean;
}

export type ChartType = "column" | "bar" | "line" | "area" | "pie" | "doughnut";
export interface ChartSeries {
  name: string;
  values: number[];
}

interface BaseObject {
  id: string;
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  stroke: string;
  roles?: ObjectRoles;
}
export interface TextObject extends BaseObject {
  kind: "text";
  text: string;
  font: string;
  size: number;
  color: string;
  align: "left" | "center" | "right";
  bold: boolean;
  italic: boolean;
}
export interface ShapeObject extends BaseObject {
  kind: "rect" | "ellipse" | "line";
}
export interface ImageObject extends BaseObject {
  kind: "image";
  resource: string;
}
export interface ChartData {
  chartType: ChartType;
  categories: string[];
  series: ChartSeries[];
  stacked: boolean;
  showLegend: boolean;
  showValues: boolean;
  unit: string;
  highlight: number | null;
}
export interface ChartObject extends BaseObject, ChartData {
  kind: "chart";
}
export type SlideObject = TextObject | ShapeObject | ImageObject | ChartObject;
export type ObjectKind = SlideObject["kind"];

export interface Slide {
  id: string;
  objects: SlideObject[];
  background?: string;
  layout?: string;
  roles?: { background: ColorRole };
  notes?: string;
  warnings?: string[];
}

export interface PresentationDocument {
  format: 1;
  id: string;
  title: string;
  width: number;
  height: number;
  theme: StoredTheme;
  slides: Slide[];
}

export interface HistoryState {
  version: number;
  document: PresentationDocument;
  undo: PresentationDocument[];
  redo: PresentationDocument[];
}

export interface Selection {
  slideId: string;
  objectIds: string[];
}

export interface Labels {
  title: string;
  subtitle: string;
  text: string;
  caption: string;
}

/** Loose record for untrusted input (tool arguments, imported data). */
export type Loose = Record<string, unknown>;
export const isRecord = (value: unknown): value is Loose =>
  typeof value === "object" && value !== null && !Array.isArray(value);
