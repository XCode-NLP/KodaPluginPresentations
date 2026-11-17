import { CHART_FIELDS, chartDefaults, validChart } from "./chart";
import { COLOR_ROLES, DEFAULT_THEME, FONT_ROLES, SIZE_ROLES, isColorRole, resolveToken, resolvedTheme, themeOf } from "./design/themes";
import {
  isRecord,
  type ColorRole,
  type HistoryState,
  type Labels,
  type Loose,
  type ObjectRoles,
  type PresentationDocument,
  type RoleField,
  type Selection,
  type Slide,
  type SlideObject,
  type Theme,
} from "./types";

export const id = () => `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 12)}`;
export const isId = (value: unknown): value is string => typeof value === "string" && /^[a-z0-9-]{1,64}$/.test(value);
const number = (value: unknown, min: number, max: number): value is number => typeof value === "number" && Number.isFinite(value) && value >= min && value <= max;
const string = (value: unknown, max: number): value is string => typeof value === "string" && value.length <= max;
const color = (value: unknown): value is string => value === "transparent" || (typeof value === "string" && /^#[\da-f]{6}$/i.test(value));
function requireValid(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}
const CHART_TYPES_TEXT = "column|bar|line|area|pie|doughnut";
const ROLE_FIELDS: Record<RoleField, readonly string[]> = { fill: COLOR_ROLES, stroke: COLOR_ROLES, color: COLOR_ROLES, font: FONT_ROLES, size: SIZE_ROLES };
const isRoleField = (field: string): field is RoleField => Object.hasOwn(ROLE_FIELDS, field);
const validRoles = (roles: unknown) =>
  roles === undefined ||
  (isRecord(roles) &&
    Object.entries(roles).every(([field, value]) => isRoleField(field) && ROLE_FIELDS[field].includes(value as string)));

function validTheme(theme: unknown): boolean {
  if (!isRecord(theme) || !color(theme.background)) return false;
  if (theme.colors === undefined && theme.fonts === undefined) return true;
  const colors = isRecord(theme.colors) ? theme.colors : {};
  const fonts = isRecord(theme.fonts) ? theme.fonts : {};
  return (
    (theme.id === null || theme.id === undefined || isId(theme.id)) &&
    COLOR_ROLES.every((role) => color(colors[role]) && colors[role] !== "transparent") &&
    FONT_ROLES.every((role) => string(fonts[role], 100) && fonts[role]) &&
    (theme.headingBold === undefined || typeof theme.headingBold === "boolean")
  );
}

/** Check an untrusted document; throws with a reason or returns it typed. */
export function validate(input: unknown): PresentationDocument {
  const doc = input as PresentationDocument;
  requireValid(
    isRecord(input) &&
      isId(doc.id) &&
      doc.format === 1 &&
      string(doc.title, 200) &&
      number(doc.width, 16, 4000) &&
      number(doc.height, 16, 4000) &&
      validTheme(doc.theme) &&
      Array.isArray(doc.slides) &&
      doc.slides.length >= 1 &&
      doc.slides.length <= 100,
    "Invalid presentation",
  );
  const ids = new Set([doc.id]);
  const unique = (key: unknown) => {
    requireValid(isId(key) && !ids.has(key), "Invalid or duplicate ID");
    ids.add(key);
  };
  for (const slide of doc.slides) {
    unique(slide.id);
    requireValid(slide.background === undefined || color(slide.background), "Invalid slide background");
    requireValid(slide.layout === undefined || isId(slide.layout), "Invalid slide layout");
    requireValid(
      slide.roles === undefined ||
        (isRecord(slide.roles) && Object.keys(slide.roles).join() === "background" && isColorRole(slide.roles.background)),
      "Invalid slide roles",
    );
    requireValid(slide.notes === undefined || string(slide.notes, 4000), "Invalid speaker notes");
    requireValid(Array.isArray(slide.objects) && slide.objects.length <= 250, "Too many slide objects");
    requireValid(
      !slide.warnings ||
        (Array.isArray(slide.warnings) &&
          slide.warnings.length <= 100 &&
          slide.warnings.every((warning) => string(warning, 300))),
      "Invalid import warnings",
    );
    for (const item of slide.objects as unknown[]) {
      requireValid(isRecord(item), "Invalid object");
      const obj = item;
      unique(obj.id);
      requireValid(
        ["text", "rect", "ellipse", "line", "image", "chart"].includes(obj.kind as string) &&
          number(obj.x, -8000, 8000) &&
          number(obj.y, -8000, 8000) &&
          number(obj.w, 1, 8000) &&
          number(obj.h, 1, 8000) &&
          color(obj.fill) &&
          color(obj.stroke),
        "Invalid object",
      );
      if (obj.kind === "text") {
        requireValid(
          string(obj.text, 12000) &&
            string(obj.font, 100) &&
            number(obj.size, 4, 600) &&
            color(obj.color) &&
            ["left", "center", "right"].includes(obj.align as string) &&
            typeof obj.bold === "boolean" &&
            typeof obj.italic === "boolean",
          "Invalid text formatting",
        );
      }
      requireValid(validRoles(obj.roles), `Invalid roles on object ${String(obj.id)}`);
      if (obj.kind === "image") requireValid(isId(obj.resource), "Invalid image resource");
      if (obj.kind === "chart")
        requireValid(
          validChart(obj),
          `Invalid chart ${String(obj.id)}: chartType ${CHART_TYPES_TEXT}; 1–24 categories (≤60 chars); 1–6 series {name, values} with one finite number per category; pie and doughnut need non-negative values`,
        );
    }
  }
  return doc;
}

export function object<T extends SlideObject = SlideObject>(kind: T["kind"], props: Partial<T> | Loose = {}): T {
  return {
    id: id(),
    kind,
    x: 80,
    y: 80,
    w: 240,
    h: 90,
    fill: kind === "text" ? "transparent" : "#acb9f5",
    stroke: "transparent",
    ...(kind === "text" && {
      text: "",
      font: "Arial",
      size: 32,
      color: "#202234",
      align: "left",
      bold: false,
      italic: false,
    }),
    ...(kind === "line" && { stroke: "#202234", h: 1 }),
    ...props,
  } as T;
}

export function slideFor(layout: string, width: number, height: number, labels: Labels, theme: Theme = resolvedTheme(DEFAULT_THEME)): Slide {
  const text = (x: number, y: number, w: number, h: number, value: string, props: Loose = {}) =>
    styled(
      object("text"),
      { x: width * x, y: height * y, w: width * w, h: height * h, text: value, font: "body", size: "body", color: "text", ...props },
      theme,
      height,
    );
  const title = text(0.067, 0.09, 0.866, 0.16, labels.title, {
    font: "heading",
    size: "h2",
    bold: theme.headingBold,
  });
  const objects: SlideObject[] = [title];
  if (layout === "title") {
    Object.assign(title, styled(title, { y: height * 0.34, h: height * 0.2, size: "h1" }, theme, height));
    objects.push(text(0.067, 0.56, 0.75, 0.14, labels.subtitle, { size: "h3", color: "muted" }));
  } else if (layout === "columns") {
    objects.push(text(0.067, 0.3, 0.41, 0.6, labels.text), text(0.522, 0.3, 0.41, 0.6, labels.text));
  } else if (layout === "caption") {
    objects.push(
      styled(object("rect"), { x: width * 0.067, y: height * 0.3, w: width * 0.866, h: height * 0.47, fill: "surface" }, theme, height),
      text(0.067, 0.8, 0.866, 0.1, labels.caption, { size: "small", color: "muted" }),
    );
  } else if (layout === "text") {
    objects.push(text(0.067, 0.3, 0.75, 0.6, labels.text));
  }
  return { id: id(), objects };
}

export function create(title: string, labels: Labels): PresentationDocument {
  return {
    format: 1,
    id: id(),
    title,
    width: 960,
    height: 540,
    theme: resolvedTheme(DEFAULT_THEME),
    slides: [slideFor("title", 960, 540, labels)],
  };
}


/** Apply a patch whose values may be role names ("accent", "h2", "heading"). */
export function styled<T extends SlideObject>(obj: T, patch: Loose, theme: Theme, height: number): T {
  const next: Loose = { ...obj };
  const roles: ObjectRoles = { ...(obj.roles ?? {}) };
  for (const [field, raw] of Object.entries(patch)) {
    if (!isRoleField(field)) {
      next[field] = raw;
      continue;
    }
    const { value, role } = resolveToken(field, raw, theme, height);
    next[field] = value;
    if (role) roles[field] = role;
    // A literal equal to the role's current value (e.g. copied from get_slide) keeps the role.
    else if (roles[field] && resolveToken(field, roles[field], theme, height).value !== value) delete roles[field];
  }
  if (Object.keys(roles).length) next.roles = roles;
  else delete next.roles;
  return next as T;
}

/** A new object may be partial: missing fields come from the theme, values may be role names. */
function addable(input: unknown, theme: Theme, height: number): SlideObject {
  requireValid(isRecord(input), "Invalid object");
  if (input.id) return styled(input as unknown as SlideObject, input, theme, height);
  requireValid(
    ["text", "rect", "ellipse", "line", "chart"].includes(input.kind as string),
    "New objects need kind text, rect, ellipse, line or chart",
  );
  const { kind, ...props } = input as Loose & { kind: SlideObject["kind"] };
  if (kind === "chart") return styled(object("chart"), { ...chartDefaults(props), ...props }, theme, height);
  const defaults =
    kind === "text"
      ? { font: "body", size: "body", color: "text" }
      : kind === "line"
        ? { stroke: "line" }
        : { fill: "surface" };
  return styled(object(kind), { ...defaults, ...props }, theme, height);
}

/** Switch the theme; objects without roles get them inferred from the old theme's values first. */
export function retheme(doc: PresentationDocument, next: Theme): void {
  const previous = themeOf(doc);
  const infer = (field: string, value: unknown): string | null => {
    if (!isRoleField(field) || typeof value !== "string") return null;
    if (field === "font") return FONT_ROLES.find((role) => previous.fonts[role] === value) ?? null;
    if (field === "size") return null;
    const order: ColorRole[] = ["text", "background", "accent", "muted", "surface", "line", "accent2", "onAccent"];
    return order.find((role) => previous.colors[role].toLowerCase() === value.toLowerCase()) ?? null;
  };
  for (const slide of doc.slides) {
    if (!slide.roles && slide.background) {
      const role = infer("fill", slide.background);
      if (role) slide.roles = { background: role as ColorRole };
    }
    if (slide.roles?.background) slide.background = next.colors[slide.roles.background];
    slide.objects = slide.objects.map((obj) => {
      const roles: ObjectRoles = { ...(obj.roles ?? {}) };
      const fields = obj as unknown as Loose;
      for (const field of ["fill", "stroke", "color", "font"] as const) {
        if (roles[field] || fields[field] === undefined || fields[field] === "transparent") continue;
        const role = infer(field, fields[field]);
        if (role) roles[field] = role;
      }
      const repainted: Loose = { ...obj };
      for (const [field, role] of Object.entries(roles)) {
        repainted[field] = resolveToken(field, role, next, doc.height).value;
      }
      if (obj.kind === "text" && roles.font === "heading" && obj.bold === previous.headingBold) repainted.bold = next.headingBold;
      if (Object.keys(roles).length) repainted.roles = roles;
      return repainted as unknown as SlideObject;
    });
  }
  doc.theme = next;
}

const SLIDE_OPERATIONS = new Set([
  "deleteSlide",
  "moveSlide",
  "updateSlide",
  "addObject",
  "deleteObject",
  "moveObject",
  "updateObject",
]);
const OPERATION_TYPES: string[] = ["rename", "setTheme", "addSlide", ...SLIDE_OPERATIONS];
const OBJECT_OPERATIONS = new Set(["deleteObject", "moveObject", "updateObject"]);

function listIds(items: { id: string }[]): string {
  const ids = items.map((item) => item.id);
  return ids.length ? ids.join(", ") : "none";
}

function indexById(items: { id: string }[], key: unknown, field = "id", owner = ""): number {
  if (key === undefined || key === null || key === "") {
    throw new Error(`Missing ${field}. Available ${field}s${owner}: ${listIds(items)}`);
  }
  const index = items.findIndex((item) => item.id === key);
  requireValid(index >= 0, `Unknown ${field} "${String(key)}". Available ${field}s${owner}: ${listIds(items)}`);
  return index;
}

// Object IDs are unique across the document, so an object operation that
// names only its objectId still points at exactly one slide.
function slideIdOfObject(doc: PresentationDocument, op: Operation): string | undefined {
  if (!OBJECT_OPERATIONS.has(op.type) || op.objectId === undefined) return undefined;
  return doc.slides.find((slide) => slide.objects.some((object) => object.id === op.objectId))?.id;
}

/** One edit as sent by the agent; fields are checked per type. */
interface Operation extends Loose {
  type: string;
  slideId?: unknown;
  objectId?: unknown;
  id?: unknown;
  index?: unknown;
  patch?: unknown;
}

function checkOperationShape(op: unknown): asserts op is Operation {
  requireValid(isRecord(op), "Operation must be an object");
  requireValid(
    OPERATION_TYPES.includes(op.type as string),
    `Unknown operation type "${String(op.type)}". Expected one of: ${OPERATION_TYPES.join(", ")}`,
  );
  if (SLIDE_OPERATIONS.has(op.type as string) && op.slideId === undefined && op.id !== undefined) {
    throw new Error(`${String(op.type)} takes "slideId", not "id"`);
  }
  if (OBJECT_OPERATIONS.has(op.type as string) && op.objectId === undefined && op.id !== undefined) {
    throw new Error(`${String(op.type)} takes "slideId" and "objectId", not "id"`);
  }
}

function reorder<T>(items: T[], index: number, to: unknown): void {
  requireValid(typeof to === "number" && Number.isInteger(to) && to >= 0 && to < items.length, "Invalid destination index");
  items.splice(to, 0, items.splice(index, 1)[0]);
}

export function apply(state: HistoryState, expectedVersion: unknown, operations: unknown): HistoryState {
  requireValid(
    expectedVersion === state.version,
    `Version conflict: current version is ${state.version}. Read the document again.`,
  );
  requireValid(
    Array.isArray(operations) && operations.length > 0 && operations.length <= 100,
    "Expected 1–100 operations",
  );
  const doc: PresentationDocument = JSON.parse(JSON.stringify(state.document));
  operations.forEach((op, position) => {
    try {
      applyOne(doc, op);
    } catch (error) {
      const fields = isRecord(op) ? Object.keys(op).join(", ") : typeof op;
      const type = isRecord(op) && op.type !== undefined ? String(op.type) : "no type";
      throw new Error(`Operation ${position} (${type}, fields: ${fields}): ${error instanceof Error ? error.message : String(error)}`);
    }
  });
  return {
    version: state.version + 1,
    document: validate(doc),
    undo: [...state.undo, state.document].slice(-50),
    redo: [],
  };
}

function applyOne(doc: PresentationDocument, op: unknown): void {
  checkOperationShape(op);
  const theme = themeOf(doc);
  {
    if (op.type === "rename") {
      doc.title = op.title as string;
      return;
    }
    if (op.type === "setTheme") {
      requireValid(typeof op.theme === "string", "setTheme needs a theme name");
      retheme(doc, resolvedTheme(op.theme, { colors: op.colors as Partial<Theme["colors"]>, fonts: op.fonts as Partial<Theme["fonts"]> }));
      return;
    }
    if (op.type === "addSlide") {
      doc.slides.push(op.slide as Slide);
      if (op.index !== undefined) reorder(doc.slides, doc.slides.length - 1, op.index);
      return;
    }
    const slideIndex = indexById(doc.slides, op.slideId ?? slideIdOfObject(doc, op), "slideId");
    const slide = doc.slides[slideIndex];
    switch (op.type) {
      case "deleteSlide":
        doc.slides.splice(slideIndex, 1);
        break;
      case "moveSlide":
        reorder(doc.slides, slideIndex, op.index);
        break;
      case "addObject":
        slide.objects.push(addable(op.object, theme, doc.height));
        break;
      case "updateSlide": {
        const patch: Loose = isRecord(op.patch) ? op.patch : {};
        requireValid(Object.keys(patch).every((field) => ["background", "notes"].includes(field)), "Invalid slide patch");
        if ("notes" in patch) {
          if (patch.notes === null || patch.notes === "") delete slide.notes;
          else slide.notes = patch.notes as string;
        }
        if ("background" in patch) {
          delete slide.roles;
          if (patch.background === null) delete slide.background;
          else {
            const { value, role } = resolveToken("fill", patch.background, theme, doc.height);
            slide.background = value as string;
            if (role) slide.roles = { background: role as ColorRole };
          }
        }
        break;
      }
      case "moveObject":
        reorder(slide.objects, indexById(slide.objects, op.objectId, "objectId", ` on slide ${slide.id}`), op.index);
        break;
      case "deleteObject":
        slide.objects.splice(indexById(slide.objects, op.objectId, "objectId", ` on slide ${slide.id}`), 1);
        break;
      case "updateObject": {
        const index = indexById(slide.objects, op.objectId, "objectId", ` on slide ${slide.id}`);
        const old = slide.objects[index];
        const fields = [
          "x",
          "y",
          "w",
          "h",
          "fill",
          "stroke",
          "text",
          "font",
          "size",
          "color",
          "align",
          "bold",
          "italic",
          "resource",
          ...(old.kind === "chart" ? CHART_FIELDS : []),
        ];
        requireValid(
          isRecord(op.patch) && Object.keys(op.patch).every((field) => fields.includes(field)),
          "Invalid object patch",
        );
        slide.objects[index] = styled(old, op.patch, theme, doc.height);
        break;
      }
      default:
        throw new Error(`Unknown operation: ${op.type}`);
    }
  }
}

export function travel(state: HistoryState, expectedVersion: unknown, direction: unknown): HistoryState {
  requireValid(
    expectedVersion === state.version,
    `Version conflict: current version is ${state.version}. Read the document again.`,
  );
  requireValid(direction === "undo" || direction === "redo", "Invalid history direction");
  const opposite = direction === "undo" ? "redo" : "undo";
  requireValid(state[direction].length, `Nothing to ${direction}`);
  return {
    ...state,
    version: state.version + 1,
    document: state[direction].at(-1)!,
    [direction]: state[direction].slice(0, -1),
    [opposite]: [...state[opposite], state.document].slice(-50),
  };
}

export function selectionFor(doc: PresentationDocument, selection: Partial<Selection> | null | undefined): Selection {
  const slide = doc.slides.find((candidate) => candidate.id === selection?.slideId) ?? doc.slides[0];
  const objectIds = slide.objects.filter((obj) => selection?.objectIds?.includes(obj.id)).map((obj) => obj.id);
  return { slideId: slide.id, objectIds };
}
