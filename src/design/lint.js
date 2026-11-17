// Layout checks the agent can act on: every issue names the slide, the object and a fix.

import { MIN_READABLE, contrast, themeOf } from "./themes.js";
import { textHeight, wordCount } from "./text.js";

const PLACEHOLDERS = new Set(["Title", "Subtitle", "Text", "Caption", "Заголовок", "Подзаголовок", "Текст", "Подпись"]);

const inside = (point, obj) => point.x >= obj.x && point.x <= obj.x + obj.w && point.y >= obj.y && point.y <= obj.y + obj.h;
const overlap = (a, b) => {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return w > 0 && h > 0 ? w * h : 0;
};

/** The color under the middle of a text box: the topmost filled shape below it, else the slide. */
function backdrop(slide, index, background) {
  const obj = slide.objects[index];
  const center = { x: obj.x + obj.w / 2, y: obj.y + Math.min(obj.h, obj.size * 1.2) / 2 };
  for (let i = index - 1; i >= 0; i--) {
    const below = slide.objects[i];
    if (below.kind === "image" && inside(center, below)) return null;
    if (["rect", "ellipse"].includes(below.kind) && below.fill !== "transparent" && inside(center, below)) return below.fill;
  }
  return background;
}

export function lintSlide(doc, slide, position) {
  const issues = [];
  const scale = doc.height / 540;
  const marginX = doc.width * 0.04;
  const marginY = doc.height * 0.04;
  const background = slide.background ?? themeOf(doc).background;
  const add = (severity, rule, message, objectId) =>
    issues.push({ slide: position + 1, slideId: slide.id, ...(objectId && { objectId }), severity, rule, message });
  let words = 0;
  const texts = [];
  slide.objects.forEach((obj, index) => {
    const right = obj.x + obj.w;
    const bottom = obj.y + obj.h;
    const bleeds = obj.kind !== "text" && (obj.x <= 0 || obj.y <= 0 || right >= doc.width || bottom >= doc.height);
    if (right <= 0 || bottom <= 0 || obj.x >= doc.width || obj.y >= doc.height)
      add("error", "offSlide", `${obj.kind} is completely outside the slide; move or delete it`, obj.id);
    else if (obj.x < 0 || obj.y < 0 || right > doc.width + 0.5 || bottom > doc.height + 0.5) {
      if (!bleeds || obj.kind === "line") add("warning", "outOfBounds", `${obj.kind} is cut off by the slide edge`, obj.id);
    }
    if (obj.kind !== "text") return;
    texts.push(obj);
    words += wordCount(obj.text);
    if (!obj.text.trim()) add("warning", "emptyText", "empty text box; delete it", obj.id);
    if (PLACEHOLDERS.has(obj.text.trim())) add("error", "placeholder", `placeholder text "${obj.text.trim()}" left on the slide`, obj.id);
    const needed = textHeight(obj.text, obj.w, obj.size, obj.font);
    if (needed > obj.h * 1.04)
      add(
        "error",
        "overflow",
        `text needs ~${Math.ceil(needed)} units of height but the box has ${Math.round(obj.h)}; shorten it, enlarge the box, or reduce size to ${Math.max(Math.round(MIN_READABLE * scale), Math.floor((obj.size * obj.h) / needed))}`,
        obj.id,
      );
    if (obj.size < MIN_READABLE * scale - 0.5)
      add("warning", "tooSmall", `font size ${obj.size} is below ${Math.round(MIN_READABLE * scale)}; cut words instead`, obj.id);
    if (obj.x < marginX - 1 || obj.y < marginY - 1 || obj.x + obj.w > doc.width - marginX + 1 || obj.y + obj.h > doc.height - marginY + 1)
      add("warning", "margin", "text is closer to the edge than the safe margin (4%)", obj.id);
    const under = backdrop(slide, index, background);
    if (under && obj.color !== "transparent" && under !== "transparent") {
      const ratio = contrast(obj.color, under);
      const large = obj.size >= 24 * scale || (obj.bold && obj.size >= 19 * scale);
      if (ratio < (large ? 3 : 4.5))
        add("error", "contrast", `contrast ${ratio.toFixed(1)}:1 between text ${obj.color} and ${under} is below ${large ? 3 : 4.5}:1`, obj.id);
    }
  });
  for (let i = 0; i < texts.length; i++) {
    for (let j = i + 1; j < texts.length; j++) {
      const area = overlap(texts[i], texts[j]);
      if (area > Math.min(texts[i].w * texts[i].h, texts[j].w * texts[j].h) * 0.08)
        add("error", "overlap", `text boxes ${texts[i].id} and ${texts[j].id} overlap`, texts[j].id);
    }
  }
  // Edges that almost, but not quite, line up read as mistakes.
  const lefts = texts.filter((obj) => obj.align === "left").map((obj) => [obj.x, obj.id]);
  for (const [x, key] of lefts) {
    const near = lefts.find(([other, otherKey]) => otherKey !== key && Math.abs(other - x) > 0.5 && Math.abs(other - x) < 8 * scale);
    if (near && x > near[0]) add("warning", "alignment", `left edge ${Math.round(x)} is ${Math.round(x - near[0])} units off ${near[1]}; align to ${Math.round(near[0])}`, key);
  }
  if (words > 70) add("warning", "density", `${words} words on one slide; aim for under 40, move details to speaker notes or split the slide`);
  const fills = new Set(
    slide.objects
      .filter((obj) => obj.kind !== "text" && obj.kind !== "image")
      .flatMap((obj) => [obj.fill, obj.stroke])
      .filter((value) => value && value !== "transparent")
      .map((value) => value.toLowerCase()),
  );
  if (fills.size > 5) add("warning", "palette", `${fills.size} different shape colors; use theme roles (surface, accent, line, chart1–5)`);
  const sizes = new Set(texts.map((obj) => Math.round(obj.size)));
  if (sizes.size > 4) add("warning", "hierarchy", `${sizes.size} different font sizes; keep to 3–4 levels from the theme scale (display, h1, h2, h3, body, small)`);
  if (!slide.objects.length) add("error", "empty", "the slide is empty");
  return issues;
}

export function lintDeck(doc, slideIds) {
  const wanted = slideIds?.length ? new Set(slideIds) : null;
  const issues = [];
  doc.slides.forEach((slide, i) => {
    if (!wanted || wanted.has(slide.id)) issues.push(...lintSlide(doc, slide, i));
  });
  if (!wanted) {
    const fonts = new Set(doc.slides.flatMap((slide) => slide.objects.filter((obj) => obj.kind === "text").map((obj) => obj.font)));
    if (fonts.size > 2) issues.push({ severity: "warning", rule: "fonts", message: `${fonts.size} font families in the deck (${[...fonts].join(", ")}); use the theme's heading and body fonts` });
    for (let i = 2; i < doc.slides.length; i++) {
      const run = [doc.slides[i - 2], doc.slides[i - 1], doc.slides[i]].map((slide) => slide.layout);
      if (run[0] && run.every((layout) => layout === run[0]))
        issues.push({ slide: i + 1, slideId: doc.slides[i].id, severity: "warning", rule: "rhythm", message: `three ${run[0]} slides in a row; vary the layout (stats, statement, image-split, cards)` });
    }
  }
  const order = { error: 0, warning: 1 };
  issues.sort((a, b) => order[a.severity] - order[b.severity]);
  return {
    errors: issues.filter((issue) => issue.severity === "error").length,
    warnings: issues.filter((issue) => issue.severity === "warning").length,
    issues: issues.slice(0, 60),
  };
}
