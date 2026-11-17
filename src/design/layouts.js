// Layout library. Each layout is drawn on a 960×540 design grid (64-unit side margins, 48-unit
// top and bottom, 8-unit rhythm) and scaled to the document. The agent fills named slots; the
// engine places, styles and fits the text, so a slide cannot drift off the grid.

import { id, object } from "../model.js";
import { MIN_READABLE, sizeFor } from "./themes.js";
import { fitText } from "./text.js";
import { CHART_TYPES, isRadial } from "../chart.js";

const M = 64;
const TOP = 48;
const W = 960;
const H = 540;
const INNER = W - M * 2;

const str = (value, max = 600) => (typeof value === "string" ? value.trim().slice(0, max) : "");
const list = (value, max) => (Array.isArray(value) ? value.slice(0, max) : []);

function painter(doc, theme, tone, images) {
  const sx = doc.width / W;
  const sy = doc.height / H;
  const objects = [];
  const dark = tone === "accent" || tone === "inverse";
  const swap = {
    accent: { text: "onAccent", muted: "onAccent", accent: "onAccent", surface: "accent", line: "onAccent" },
    inverse: { text: "background", muted: "line", surface: "text", line: "muted", background: "text" },
  }[tone] ?? {};
  const role = (name) => swap[name] ?? name;
  const box = (x, y, w, h) => ({ x: x * sx, y: y * sy, w: Math.max(1, w * sx), h: Math.max(1, h * sy) });
  const paint = (roles) => {
    const values = {};
    for (const [field, value] of Object.entries(roles)) {
      if (field === "size") values.size = sizeFor(value, doc.height);
      else if (field === "font") values.font = theme.fonts[value];
      else values[field] = value === "transparent" ? "transparent" : theme.colors[value];
    }
    return values;
  };
  const clean = (roles) => Object.fromEntries(Object.entries(roles).filter(([, value]) => value !== "transparent"));
  return {
    objects,
    dark,
    background: { slide: dark ? (tone === "accent" ? "accent" : "text") : null },
    text(value, x, y, w, h, { size = "body", font = "body", color = "text", align = "left", bold, italic = false, min } = {}) {
      value = str(value, 2000);
      if (!value) return null;
      const roles = { size, font, color: role(color) };
      const placed = box(x, y, w, h);
      const styled = paint(roles);
      const floor = Math.max(sizeFor("caption", doc.height), ((min ?? MIN_READABLE) * doc.height) / 540);
      const fit = fitText(value, placed.w, placed.h, styled.size, styled.font, floor);
      const obj = object("text", {
        ...placed,
        ...styled,
        text: value,
        size: fit.size,
        align,
        italic,
        bold: bold ?? (font === "heading" ? theme.headingBold : false),
        roles: fit.size === styled.size ? roles : { font: roles.font, color: roles.color },
      });
      objects.push(obj);
      return obj;
    },
    shape(kind, x, y, w, h, fill = "surface", stroke = "transparent") {
      const roles = clean({ fill: role(fill), stroke: role(stroke) });
      const obj = object(kind, {
        ...box(x, y, w, h),
        fill: "transparent",
        stroke: "transparent",
        ...paint(roles),
        roles,
      });
      objects.push(obj);
      return obj;
    },
    line(x, y, w, color = "line") {
      const roles = { stroke: role(color) };
      const obj = object("line", { ...box(x, y, w, 1), h: 1, fill: "transparent", ...paint(roles), roles });
      objects.push(obj);
      return obj;
    },
    chart(x, y, w, h, data) {
      const obj = object("chart", { ...box(x, y, w, h), fill: "transparent", stroke: "transparent", ...data });
      objects.push(obj);
      return obj;
    },
    /** Fit an image inside the box without cropping; without an image draw a quiet placeholder block. */
    image(slot, x, y, w, h, { placeholder = true } = {}) {
      const image = images[slot];
      if (!image) return placeholder ? this.shape("rect", x, y, w, h, "surface") : null;
      const placed = box(x, y, w, h);
      const scale = Math.min(placed.w / image.width, placed.h / image.height);
      const iw = image.width * scale;
      const ih = image.height * scale;
      const obj = object("image", {
        x: placed.x + (placed.w - iw) / 2,
        y: placed.y + (placed.h - ih) / 2,
        w: Math.max(1, iw),
        h: Math.max(1, ih),
        fill: "transparent",
        stroke: "transparent",
        resource: image.resource,
      });
      objects.push(obj);
      return obj;
    },
  };
}

function title(p, value, { y = TOP, h = 90, w = INNER, size = "h2" } = {}) {
  return p.text(value, M, y, w, h, { size, font: "heading" });
}

/** Columns of equal width across the content area. */
const columns = (count, gap = 24, left = M, width = INNER) => {
  const w = (width - gap * (count - 1)) / count;
  return Array.from({ length: count }, (_, i) => ({ x: left + i * (w + gap), w }));
};

const LAYOUTS = {
  title: {
    use: "Opening slide: deck title, one-line promise, author/date",
    slots: { kicker: "text ≤5 words, optional", title: "text ≤10 words", subtitle: "text ≤20 words", footer: "text ≤10 words, author · date" },
    build(p, s) {
      p.shape("rect", M, 150, 56, 6, "accent");
      p.text(s.kicker, M, 112, INNER, 30, { size: "small", color: "accent", bold: true });
      p.text(s.title, M, 176, 780, 170, { size: "display", font: "heading" });
      p.text(s.subtitle, M, 356, 700, 70, { size: "h3", color: "muted" });
      p.text(s.footer, M, 456, INNER, 32, { size: "small", color: "muted" });
    },
  },
  section: {
    use: "Chapter divider between parts of the story",
    tone: "accent",
    slots: { number: "text like 02, optional", title: "text ≤6 words", subtitle: "text ≤16 words, optional" },
    build(p, s) {
      p.text(s.number, M, 120, 300, 90, { size: "display", font: "heading", color: "muted" });
      p.text(s.title, M, 220, 800, 130, { size: "h1", font: "heading" });
      p.text(s.subtitle, M, 360, 700, 70, { size: "body", color: "muted" });
    },
  },
  statement: {
    use: "One big idea or key takeaway in a single sentence",
    slots: { statement: "text ≤20 words", support: "text ≤20 words, optional" },
    build(p, s) {
      p.shape("rect", M, 130, 56, 6, "accent");
      p.text(s.statement, M, 156, 800, 200, { size: "h1", font: "heading" });
      p.text(s.support, M, 380, 640, 80, { size: "h3", color: "muted" });
    },
  },
  agenda: {
    use: "Table of contents, 3–6 items",
    slots: { title: "text", items: "list of 3–6 short strings" },
    build(p, s) {
      title(p, s.title);
      const items = list(s.items, 6);
      const row = Math.min(60, 330 / Math.max(1, items.length));
      items.forEach((item, i) => {
        const y = 156 + i * row;
        p.text(String(i + 1).padStart(2, "0"), M, y + 8, 60, row - 12, { size: "h3", font: "heading", color: "accent", bold: true });
        p.text(item, M + 76, y + 10, INNER - 76, row - 14, { size: "h3" });
        p.line(M, y + row - 1, INNER);
      });
    },
  },
  bullets: {
    use: "Title plus 3–5 short points, optional side note",
    slots: { title: "text ≤10 words, phrase it as the conclusion", points: "list of 3–5 strings ≤12 words", note: "text ≤30 words, optional" },
    build(p, s) {
      title(p, s.title);
      const points = list(s.points, 6).map((item) => str(item)).filter(Boolean);
      const width = s.note ? 540 : INNER;
      p.text(points.map((point) => `•  ${point}`).join("\n\n"), M, 164, width, 320, { size: "body" });
      if (s.note) {
        p.shape("rect", 640, 164, 256, 300, "surface");
        p.shape("rect", 640, 164, 4, 300, "accent");
        p.text(s.note, 664, 188, 212, 252, { size: "small", color: "muted" });
      }
    },
  },
  text: {
    use: "Title plus one paragraph; prefer bullets or statement when possible",
    slots: { title: "text", body: "text ≤60 words" },
    build(p, s) {
      title(p, s.title);
      p.text(s.body, M, 164, 680, 320, { size: "body" });
    },
  },
  "two-columns": {
    use: "Two parallel ideas side by side",
    slots: { title: "text", columns: "list of 2 {title, text ≤35 words}" },
    build(p, s) {
      title(p, s.title);
      const cols = columns(2, 48);
      list(s.columns, 2).forEach((col, i) => {
        p.shape("rect", cols[i].x, 164, 40, 4, i ? "accent2" : "accent");
        p.text(col?.title, cols[i].x, 180, cols[i].w, 70, { size: "h3", font: "heading", bold: true });
        p.text(col?.text, cols[i].x, 256, cols[i].w, 230, { size: "body", color: "muted" });
      });
    },
  },
  cards: {
    use: "2–4 features, options, pillars or benefits as cards",
    slots: { title: "text", cards: "list of 2–4 {title ≤4 words, text ≤20 words, label optional, e.g. 01}" },
    build(p, s) {
      title(p, s.title);
      const cards = list(s.cards, 4);
      const cols = columns(Math.max(1, cards.length));
      cards.forEach((card, i) => {
        const { x, w } = cols[i];
        p.shape("rect", x, 164, w, 316, "surface");
        p.text(card?.label ?? String(i + 1).padStart(2, "0"), x + 24, 188, w - 48, 30, { size: "small", color: "accent", bold: true });
        p.text(card?.title, x + 24, 228, w - 48, 80, { size: "h3", font: "heading", bold: true });
        p.text(card?.text, x + 24, 316, w - 48, 150, { size: "small", color: "muted" });
      });
    },
  },
  "big-number": {
    use: "One striking metric with its meaning",
    slots: { value: "text like 3.2× or 87%", label: "text ≤8 words, what the number measures", context: "text ≤25 words, why it matters, optional" },
    build(p, s) {
      p.text(s.value, M, 96, INNER, 180, { size: "display", font: "heading", color: "accent", bold: true });
      const big = p.objects.at(-1);
      if (big) {
        const target = Math.round((150 * big.h) / 180);
        if (big.size < target && fitText(big.text, big.w, big.h, target, big.font, target).fits) {
          big.size = target;
          delete big.roles.size;
        }
      }
      p.text(s.label, M, 290, 760, 70, { size: "h2", font: "heading" });
      p.text(s.context, M, 376, 640, 90, { size: "body", color: "muted" });
    },
  },
  stats: {
    use: "2–4 key metrics in a row",
    slots: { title: "text", stats: "list of 2–4 {value like 42%, label ≤8 words}", note: "text ≤20 words, optional source or takeaway" },
    build(p, s) {
      title(p, s.title);
      const stats = list(s.stats, 4);
      const cols = columns(Math.max(1, stats.length), 32);
      stats.forEach((stat, i) => {
        const { x, w } = cols[i];
        p.line(x, 196, w, i ? "line" : "accent");
        p.text(stat?.value, x, 216, w, 90, { size: "h1", font: "heading", color: i ? "text" : "accent", bold: true });
        p.text(stat?.label, x, 312, w, 90, { size: "body", color: "muted" });
      });
      p.text(s.note, M, 452, INNER, 40, { size: "caption", color: "muted" });
    },
  },
  comparison: {
    use: "Before/after, us vs them, option A vs B, pros and cons",
    slots: { title: "text", left: "{title, points: list of 2–5 strings}", right: "{title, points: list of 2–5 strings}; right is the recommended side" },
    build(p, s) {
      title(p, s.title);
      const cols = columns(2, 32);
      [s.left, s.right].forEach((side, i) => {
        const { x, w } = cols[i];
        p.shape("rect", x, 164, w, 316, i ? "surface" : "background", i ? "transparent" : "line");
        p.shape("rect", x, 164, w, 6, i ? "accent" : "muted");
        p.text(side?.title, x + 28, 190, w - 56, 50, { size: "h3", font: "heading", bold: true, color: i ? "accent" : "text" });
        const points = list(side?.points, 5).map((item) => str(item)).filter(Boolean);
        p.text(points.map((point) => `${i ? "✓" : "–"}  ${point}`).join("\n"), x + 28, 252, w - 56, 210, { size: "body", color: i ? "text" : "muted" });
      });
    },
  },
  timeline: {
    use: "Roadmap or history, 3–6 milestones in order",
    slots: { title: "text", steps: "list of 3–6 {label like Q1 or 2025, title ≤4 words, text ≤12 words}" },
    build(p, s) {
      title(p, s.title);
      const steps = list(s.steps, 6);
      const cols = columns(Math.max(1, steps.length), 20);
      p.line(M, 236, INNER);
      steps.forEach((step, i) => {
        const { x, w } = cols[i];
        p.text(step?.label, x, 180, w, 36, { size: "small", color: "accent", bold: true });
        p.shape("ellipse", x, 228, 16, 16, i === steps.length - 1 ? "accent" : "text");
        p.text(step?.title, x, 268, w, 64, { size: "h3", font: "heading", bold: true });
        p.text(step?.text, x, 338, w - 8, 140, { size: "small", color: "muted" });
      });
    },
  },
  process: {
    use: "How something works: 3–5 sequential steps",
    slots: { title: "text", steps: "list of 3–5 {title ≤4 words, text ≤15 words}" },
    build(p, s) {
      title(p, s.title);
      const steps = list(s.steps, 5);
      const cols = columns(Math.max(1, steps.length), 40);
      steps.forEach((step, i) => {
        const { x, w } = cols[i];
        p.shape("ellipse", x, 172, 56, 56, "accent");
        p.text(String(i + 1), x, 184, 56, 36, { size: "h3", font: "heading", color: "onAccent", align: "center", bold: true });
        if (i < steps.length - 1) p.line(x + 72, 200, w - 72 + 24, "line");
        p.text(step?.title, x, 252, w, 64, { size: "h3", font: "heading", bold: true });
        p.text(step?.text, x, 322, w, 150, { size: "small", color: "muted" });
      });
    },
  },
  quote: {
    use: "Customer or expert quote",
    slots: { quote: "text ≤35 words", author: "text, name", role: "text, title and company, optional" },
    build(p, s) {
      p.text("“", M, 48, 120, 110, { size: "display", font: "heading", color: "accent", bold: true });
      p.text(s.quote, M, 168, 800, 212, { size: "h2", font: "heading", italic: true });
      p.shape("rect", M, 404, 40, 4, "accent");
      p.text(s.author, M, 420, 600, 34, { size: "body", bold: true });
      p.text(s.role, M, 454, 600, 30, { size: "small", color: "muted" });
    },
  },
  "image-split": {
    use: "Screenshot, product shot or photo next to a short explanation",
    slots: { image: "project PNG path", title: "text ≤8 words", body: "text ≤40 words or list of up to 4 points", side: "image side: left (default) or right" },
    build(p, s) {
      const right = s.side === "right";
      const imageX = right ? 480 : 0;
      const textX = right ? M : 480 + 48;
      p.shape("rect", imageX, 0, 480, 540, "surface");
      p.image("image", imageX + 32, 32, 416, 476, { placeholder: false });
      p.text(s.title, textX, 120, 368, 120, { size: "h2", font: "heading" });
      const body = Array.isArray(s.body) ? list(s.body, 4).map((point) => `•  ${str(point)}`).join("\n\n") : s.body;
      p.text(body, textX, 256, 368, 230, { size: "body", color: "muted" });
    },
  },
  "image-full": {
    use: "Full-slide visual with a caption band",
    tone: "inverse",
    slots: { image: "project PNG path", caption: "text ≤15 words" },
    build(p, s) {
      p.image("image", 0, 0, 960, 448, { placeholder: false });
      p.shape("rect", 0, 448, 960, 92, "surface");
      p.text(s.caption, M, 470, INNER, 50, { size: "h3", font: "heading", color: "text" });
    },
  },
  chart: {
    use: "Native chart with the takeaway beside it: column, bar, line, area, pie or doughnut",
    slots: {
      title: "text, the conclusion the data shows",
      type: "column (default), bar, line, area, pie or doughnut",
      data: "one series: list of 2–12 {label ≤2 words, value number}",
      categories: "several series instead of data: list of 2–12 labels",
      series: "several series instead of data: list of 1–5 {name, values: one number per category}",
      stacked: "true to stack the series (column, bar, area)",
      unit: "text like %, ms, $M, optional",
      highlight: "index of the bar to emphasize with one series, default the largest",
      insight: "text ≤25 words, the takeaway",
    },
    build(p, s) {
      title(p, s.title);
      const type = CHART_TYPES.includes(s.type) ? s.type : "column";
      let categories;
      let series;
      if (Array.isArray(s.series) && s.series.length) {
        categories = list(s.categories, 12).map((item) => str(item, 60));
        series = list(s.series, 5)
          .map((item) => ({ name: str(item?.name, 60), values: list(item?.values, categories.length).map(Number) }))
          .filter((item) => item.values.length === categories.length && item.values.every(Number.isFinite));
      } else {
        const data = list(s.data, 12).filter((item) => Number.isFinite(item?.value));
        categories = data.map((item) => str(item.label, 60));
        series = [{ name: str(s.title, 60), values: data.map((item) => item.value) }];
      }
      if (!categories.length || !series.length) return;
      if (isRadial(type)) series = series.slice(0, 1).map((item) => ({ ...item, values: item.values.map((v) => Math.max(0, v)) }));
      let highlight = null;
      if (series.length === 1 && (type === "column" || type === "bar")) {
        const values = series[0].values;
        const max = Math.max(...values.map(Math.abs));
        highlight = Number.isInteger(s.highlight) && s.highlight >= 0 && s.highlight < values.length ? s.highlight : values.findIndex((v) => Math.abs(v) === max);
      }
      const width = s.insight ? 568 : INNER;
      p.chart(M, 150, width, 340, {
        chartType: type,
        categories,
        series,
        stacked: s.stacked === true && !isRadial(type) && type !== "line",
        showLegend: series.length > 1 || isRadial(type),
        showValues: series.length === 1,
        unit: str(s.unit, 12),
        highlight,
      });
      if (s.insight) {
        p.shape("rect", 664, 164, 232, 316, "surface");
        p.shape("rect", 664, 164, 4, 316, "accent");
        p.text(s.insight, 688, 188, 188, 270, { size: "h3", font: "heading" });
      }
    },
  },
  table: {
    use: "Small table, up to 5 columns and 6 rows; for more, split the slide",
    slots: { title: "text", columns: "list of 2–5 header strings", rows: "list of up to 6 rows, each a list of cell strings", highlight: "index of the column to emphasize, optional" },
    build(p, s) {
      title(p, s.title);
      const headers = list(s.columns, 5).map((item) => str(item));
      const rows = list(s.rows, 6);
      const cols = columns(Math.max(1, headers.length), 16);
      const rowH = Math.min(48, 260 / Math.max(1, rows.length));
      if (Number.isInteger(s.highlight) && cols[s.highlight]) {
        const { x, w } = cols[s.highlight];
        p.shape("rect", x - 8, 160, w + 16, 52 + rows.length * rowH + 8, "surface");
      }
      headers.forEach((header, i) => p.text(header, cols[i].x, 168, cols[i].w, 36, { size: "small", bold: true, color: "accent" }));
      p.line(M, 208, INNER, "text");
      rows.forEach((row, r) => {
        const y = 216 + r * rowH;
        list(row, headers.length).forEach((cell, i) => p.text(String(cell ?? ""), cols[i].x, y + 6, cols[i].w, rowH - 8, { size: "small", min: 11 }));
        p.line(M, y + rowH, INNER);
      });
    },
  },
  team: {
    use: "People: team, speakers, contacts, up to 4",
    slots: { title: "text", people: "list of 1–4 {name, role ≤6 words, photo: project PNG path, optional}" },
    build(p, s, images) {
      title(p, s.title);
      const people = list(s.people, 4);
      const cols = columns(Math.max(1, people.length), 32);
      people.forEach((person, i) => {
        const { x, w } = cols[i];
        const size = Math.min(150, w);
        const left = x + (w - size) / 2;
        if (images[`people.${i}.photo`]) p.image(`people.${i}.photo`, left, 168, size, size);
        else {
          p.shape("ellipse", left, 168, size, size, "surface");
          const initials = str(person?.name).split(/\s+/).map((part) => part[0] ?? "").join("").slice(0, 2).toUpperCase();
          p.text(initials, left, 168 + size / 2 - 22, size, 44, { size: "h2", font: "heading", color: "accent", align: "center", bold: true });
        }
        p.text(person?.name, x, 340, w, 40, { size: "h3", font: "heading", align: "center", bold: true });
        p.text(person?.role, x, 384, w, 60, { size: "small", color: "muted", align: "center" });
      });
    },
  },
  closing: {
    use: "Final slide: call to action, thanks, contacts",
    tone: "accent",
    slots: { title: "text ≤8 words, the ask or thanks", subtitle: "text ≤20 words, optional", contacts: "text ≤15 words, optional" },
    build(p, s) {
      p.text(s.title, M, 160, INNER, 150, { size: "h1", font: "heading", align: "center" });
      p.text(s.subtitle, 160, 320, 640, 80, { size: "h3", align: "center", color: "muted" });
      p.text(s.contacts, M, 446, INNER, 40, { size: "small", align: "center", color: "muted" });
    },
  },
};

export const LAYOUT_IDS = Object.keys(LAYOUTS);

export const layoutCatalog = () =>
  Object.entries(LAYOUTS).map(([key, { use, slots, tone }]) => ({ id: key, use, slots, ...(tone && { tone }) }));

/** Slots that name project PNG files: "image", "people.0.photo". */
export function imageSlots(layout, slots) {
  if (layout === "image-split" || layout === "image-full") return typeof slots?.image === "string" ? { image: slots.image } : {};
  if (layout === "team")
    return Object.fromEntries(
      list(slots?.people, 4).flatMap((person, i) => (typeof person?.photo === "string" ? [[`people.${i}.photo`, person.photo]] : [])),
    );
  return {};
}

/**
 * Build a slide. `images` maps slot keys to { resource, width, height } already imported into the
 * document. `tone` overrides the layout default: "default", "accent" or "inverse".
 */
export function buildLayoutSlide(doc, theme, layout, slots = {}, { images = {}, tone } = {}) {
  const spec = LAYOUTS[layout];
  if (!spec) throw new Error(`Unknown layout: ${layout}. Call get_design_system for the list.`);
  const chosen = tone && tone !== "default" ? tone : tone === "default" ? null : spec.tone;
  const p = painter(doc, theme, chosen, images);
  spec.build(p, slots ?? {}, images);
  if (!p.objects.length) throw new Error(`Layout ${layout} got no content: fill its slots`);
  const slide = { id: id(), layout, objects: p.objects };
  if (chosen === "accent" || chosen === "inverse") {
    const role = chosen === "accent" ? "accent" : "text";
    slide.background = theme.colors[role];
    slide.roles = { background: role };
  }
  return slide;
}
