// Native chart objects. A chart stores only data; its colors come from the theme's chart roles at
// render time, so a theme switch repaints it. The panel draws it as SVG, PowerPoint gets a real,
// editable chart.

export const CHART_TYPES = ["column", "bar", "line", "area", "pie", "doughnut"];
export const CHART_FIELDS = ["chartType", "categories", "series", "stacked", "showLegend", "showValues", "unit", "highlight"];

const MAX_CATEGORIES = 24;
const MAX_SERIES = 6;
const FALLBACK = ["#3d5afe", "#00796f", "#f5a524", "#e5484d", "#8e4ec6", "#5d6276"];

const finite = (value) => typeof value === "number" && Number.isFinite(value) && Math.abs(value) < 1e12;
const label = (value, max) => typeof value === "string" && value.length <= max;
const round = (value) => Number(value.toFixed(2));
export const isRadial = (type) => type === "pie" || type === "doughnut";

export function validChart(obj) {
  const n = Array.isArray(obj.categories) ? obj.categories.length : 0;
  return (
    CHART_TYPES.includes(obj.chartType) &&
    n >= 1 &&
    n <= MAX_CATEGORIES &&
    obj.categories.every((item) => label(item, 60)) &&
    Array.isArray(obj.series) &&
    obj.series.length >= 1 &&
    obj.series.length <= MAX_SERIES &&
    obj.series.every(
      (s) => s && label(s.name, 60) && Array.isArray(s.values) && s.values.length === n && s.values.every(finite),
    ) &&
    (!isRadial(obj.chartType) || obj.series[0].values.every((v) => v >= 0)) &&
    typeof obj.stacked === "boolean" &&
    typeof obj.showLegend === "boolean" &&
    typeof obj.showValues === "boolean" &&
    label(obj.unit, 12) &&
    (obj.highlight === null || (Number.isInteger(obj.highlight) && obj.highlight >= 0 && obj.highlight < n))
  );
}

/** Defaults for a new chart: one series reads without a legend, several need one. */
export function chartDefaults(input) {
  const series = Array.isArray(input.series) ? input.series : [];
  const type = input.chartType ?? "column";
  return {
    chartType: type,
    stacked: false,
    showLegend: series.length > 1 || isRadial(type),
    showValues: series.length === 1,
    unit: "",
    highlight: null,
    fill: "transparent",
  };
}

export function chartPalette(theme) {
  const c = theme?.colors;
  return {
    series: c ? [c.chart1, c.chart2, c.chart3, c.chart4, c.chart5, c.muted] : FALLBACK,
    text: c?.text ?? "#1c1f2e",
    muted: c?.muted ?? "#5d6276",
    line: c?.line ?? "#dfe2ea",
    accent: c?.accent ?? FALLBACK[0],
    background: theme?.background ?? "#ffffff",
    font: theme?.fonts?.body ?? "Arial",
  };
}

/** Per-category colors for a single highlighted series, else null. */
function pointColors(obj, palette) {
  if (obj.series.length !== 1 || obj.highlight === null || !["column", "bar"].includes(obj.chartType)) return null;
  return obj.categories.map((_, i) => (i === obj.highlight ? palette.accent : palette.line));
}

function niceStep(span) {
  const raw = span / 5;
  const power = 10 ** Math.floor(Math.log10(raw || 1));
  return [1, 2, 2.5, 5, 10].map((m) => m * power).find((step) => step >= raw) ?? power * 10;
}

function scale(obj) {
  const { series, categories, stacked } = obj;
  let min = 0;
  let max = 0;
  categories.forEach((_, i) => {
    if (stacked) {
      const pos = series.reduce((sum, s) => sum + Math.max(0, s.values[i]), 0);
      const neg = series.reduce((sum, s) => sum + Math.min(0, s.values[i]), 0);
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    } else
      for (const s of series) {
        max = Math.max(max, s.values[i]);
        min = Math.min(min, s.values[i]);
      }
  });
  if (max === min) max = min + 1;
  const step = niceStep(max - min);
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(round(v));
  return { lo, hi, ticks };
}

const escape = (value) =>
  String(value).replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch]);
const format = (value, unit) => `${round(value).toLocaleString("ru-RU")}${unit}`;

/** SVG markup for a chart object, sized to its box. Text is escaped; no external references. */
export function chartSvg(obj, theme) {
  const p = chartPalette(theme);
  const { w, h } = obj;
  const fs = Math.max(10, Math.min(18, Math.min(w, h) / 22));
  const out = [];
  const text = (x, y, value, { anchor = "middle", color = p.muted, weight = "normal", size = fs, baseline = "middle" } = {}) =>
    out.push(
      `<text x="${round(x)}" y="${round(y)}" text-anchor="${anchor}" dominant-baseline="${baseline}" fill="${color}" font-size="${round(size)}" font-weight="${weight}">${escape(value)}</text>`,
    );
  const radial = isRadial(obj.chartType);
  const legendItems = radial ? obj.categories : obj.series.map((s) => s.name);
  const colorOf = (i) => p.series[i % p.series.length];
  let top = 4;
  if (obj.showLegend) {
    const gap = fs * 1.2;
    const widths = legendItems.map((name) => fs * 1.2 + name.length * fs * 0.55 + gap);
    let x = Math.max(0, (w - widths.reduce((a, b) => a + b, 0)) / 2);
    let y = fs * 0.8;
    legendItems.forEach((name, i) => {
      if (x + widths[i] > w && x > 0) {
        x = 0;
        y += fs * 1.6;
      }
      out.push(`<rect x="${round(x)}" y="${round(y - fs * 0.4)}" width="${round(fs * 0.8)}" height="${round(fs * 0.8)}" rx="2" fill="${colorOf(i)}"/>`);
      text(x + fs * 1.1, y, name, { anchor: "start", color: p.text });
      x += widths[i];
    });
    top = y + fs * 1.4;
  }
  if (radial) {
    const values = obj.series[0].values;
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const r = Math.max(4, Math.min(w, h - top) / 2 - fs * 0.5);
    const cx = w / 2;
    const cy = top + (h - top) / 2;
    const inner = obj.chartType === "doughnut" ? r * 0.55 : 0;
    let angle = -Math.PI / 2;
    values.forEach((value, i) => {
      if (!value) return;
      const sweep = (value / total) * Math.PI * 2;
      const end = angle + Math.min(sweep, Math.PI * 2 - 1e-4);
      const large = sweep > Math.PI ? 1 : 0;
      const pt = (radius, a) => `${round(cx + radius * Math.cos(a))} ${round(cy + radius * Math.sin(a))}`;
      const path = inner
        ? `M${pt(r, angle)} A${round(r)} ${round(r)} 0 ${large} 1 ${pt(r, end)} L${pt(inner, end)} A${round(inner)} ${round(inner)} 0 ${large} 0 ${pt(inner, angle)}Z`
        : `M${round(cx)} ${round(cy)} L${pt(r, angle)} A${round(r)} ${round(r)} 0 ${large} 1 ${pt(r, end)}Z`;
      out.push(`<path d="${path}" fill="${colorOf(i)}" stroke="${p.background}" stroke-width="2"/>`);
      if (obj.showValues && value / total >= 0.04) {
        const mid = angle + sweep / 2;
        const lr = inner ? (r + inner) / 2 : r * 0.62;
        text(cx + lr * Math.cos(mid), cy + lr * Math.sin(mid), `${Math.round((value / total) * 100)}%`, { color: "#ffffff", weight: "bold" });
      }
      angle += sweep;
    });
  } else {
    const { lo, hi, ticks } = scale(obj);
    const horizontal = obj.chartType === "bar";
    const tickLabels = ticks.map((v) => format(v, obj.unit));
    const longest = (items) => Math.max(...items.map((s) => s.length)) * fs * 0.55;
    const left = Math.min(w * 0.4, (horizontal ? longest(obj.categories) : longest(tickLabels)) + fs * 0.8);
    const bottom = fs * 2;
    const plot = { x: left, y: top + fs * (obj.showValues ? 1.2 : 0.4), w: Math.max(4, w - left - fs * 0.6), h: 0 };
    plot.h = Math.max(4, h - plot.y - bottom);
    const n = obj.categories.length;
    const valuePos = (v) => (horizontal ? plot.x + ((v - lo) / (hi - lo)) * plot.w : plot.y + plot.h - ((v - lo) / (hi - lo)) * plot.h);
    // Grid and value axis labels.
    ticks.forEach((v, i) => {
      const at = valuePos(v);
      if (horizontal) {
        out.push(`<line x1="${round(at)}" y1="${round(plot.y)}" x2="${round(at)}" y2="${round(plot.y + plot.h)}" stroke="${p.line}" stroke-width="1"/>`);
        text(at, plot.y + plot.h + fs, tickLabels[i]);
      } else {
        out.push(`<line x1="${round(plot.x)}" y1="${round(at)}" x2="${round(plot.x + plot.w)}" y2="${round(at)}" stroke="${p.line}" stroke-width="1"/>`);
        text(plot.x - fs * 0.5, at, tickLabels[i], { anchor: "end" });
      }
    });
    const band = (horizontal ? plot.h : plot.w) / n;
    const bandStart = (i) => (horizontal ? plot.y : plot.x) + band * i;
    obj.categories.forEach((name, i) => {
      if (horizontal) text(plot.x - fs * 0.5, bandStart(i) + band / 2, name, { anchor: "end", color: p.text });
      else if (n <= 12 || i % Math.ceil(n / 12) === 0) text(bandStart(i) + band / 2, plot.y + plot.h + fs, name, { color: p.text });
    });
    const zero = valuePos(Math.max(lo, Math.min(hi, 0)));
    const points = pointColors(obj, p);
    if (obj.chartType === "column" || obj.chartType === "bar") {
      const groups = obj.stacked ? 1 : obj.series.length;
      const inner = band * 0.72;
      const thick = inner / groups;
      const stackPos = obj.categories.map(() => 0);
      const stackNeg = obj.categories.map(() => 0);
      obj.series.forEach((s, si) => {
        s.values.forEach((v, i) => {
          const base = obj.stacked ? (v >= 0 ? stackPos[i] : stackNeg[i]) : 0;
          if (obj.stacked) v >= 0 ? (stackPos[i] += v) : (stackNeg[i] += v);
          const a = valuePos(base);
          const b = valuePos(base + v);
          const offset = bandStart(i) + (band - inner) / 2 + (obj.stacked ? 0 : si * thick);
          const fill = points ? points[i] : colorOf(si);
          const [x, y, bw, bh] = horizontal
            ? [Math.min(a, b), offset, Math.abs(b - a), thick - 1]
            : [offset, Math.min(a, b), thick - 1, Math.abs(b - a)];
          out.push(`<rect x="${round(x)}" y="${round(y)}" width="${round(Math.max(1, bw))}" height="${round(Math.max(1, bh))}" fill="${fill}"/>`);
          if (obj.showValues && !obj.stacked) {
            const strong = points && i === obj.highlight;
            const color = strong ? p.accent : p.muted;
            const weight = strong ? "bold" : "normal";
            if (horizontal) text(b + (v >= 0 ? fs * 0.4 : -fs * 0.4), offset + thick / 2, format(v, obj.unit), { anchor: v >= 0 ? "start" : "end", color, weight });
            else text(offset + thick / 2, b + (v >= 0 ? -fs * 0.7 : fs * 0.7), format(v, obj.unit), { color, weight });
          }
        });
      });
    } else {
      const cumulative = obj.categories.map(() => 0);
      const layers = obj.series.map((s) =>
        s.values.map((v, i) => {
          const base = obj.stacked ? cumulative[i] : 0;
          if (obj.stacked) cumulative[i] += v;
          return { x: bandStart(i) + band / 2, y: valuePos(base + v), base: valuePos(base), v };
        }),
      );
      layers.forEach((pts, si) => {
        const color = colorOf(si);
        const line = pts.map((pt, i) => `${i ? "L" : "M"}${round(pt.x)} ${round(pt.y)}`).join(" ");
        if (obj.chartType === "area") {
          const back = [...pts].reverse().map((pt) => `L${round(pt.x)} ${round(obj.stacked ? pt.base : zero)}`).join(" ");
          out.push(`<path d="${line} ${back}Z" fill="${color}" fill-opacity="${obj.stacked ? 0.85 : 0.35}"/>`);
        }
        out.push(`<path d="${line}" fill="none" stroke="${color}" stroke-width="${round(Math.max(2, fs / 6))}" stroke-linejoin="round"/>`);
        if (obj.chartType === "line")
          for (const pt of pts) out.push(`<circle cx="${round(pt.x)}" cy="${round(pt.y)}" r="${round(Math.max(2.5, fs / 4.5))}" fill="${color}"/>`);
        if (obj.showValues) for (const pt of pts) text(pt.x, pt.y - fs * 0.8, format(pt.v, obj.unit), { color: p.text });
      });
    }
    if (horizontal)
      out.push(`<line x1="${round(zero)}" y1="${round(plot.y)}" x2="${round(zero)}" y2="${round(plot.y + plot.h)}" stroke="${p.muted}" stroke-width="1"/>`);
    else out.push(`<line x1="${round(plot.x)}" y1="${round(zero)}" x2="${round(plot.x + plot.w)}" y2="${round(zero)}" stroke="${p.muted}" stroke-width="1"/>`);
  }
  const font = escape(p.font.replace(/["\\]/g, ""));
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${round(w)}" height="${round(h)}" viewBox="0 0 ${round(w)} ${round(h)}" font-family="${font}, sans-serif">${out.join("")}</svg>`;
}

/** pptxgenjs chart type and options for a chart object (box in inches). */
export function pptxChart(pptx, obj, theme, box) {
  const p = chartPalette(theme);
  const hex = (value) => value.slice(1);
  const type = {
    column: pptx.ChartType.bar,
    bar: pptx.ChartType.bar,
    line: pptx.ChartType.line,
    area: pptx.ChartType.area,
    pie: pptx.ChartType.pie,
    doughnut: pptx.ChartType.doughnut,
  }[obj.chartType];
  const radial = isRadial(obj.chartType);
  const series = radial ? obj.series.slice(0, 1) : obj.series;
  const data = series.map((s) => ({ name: s.name, labels: obj.categories, values: s.values }));
  const points = pointColors(obj, p);
  const fontSize = Math.max(8, Math.min(14, Math.min(obj.w, obj.h) / 30));
  const unit = obj.unit.replace(/"/g, "");
  const numberFormat = unit ? `#,##0.##"${unit}"` : "#,##0.##";
  const options = {
    ...box,
    chartColors: (points ?? (radial ? obj.categories.map((_, i) => p.series[i % p.series.length]) : p.series)).map(hex),
    showLegend: obj.showLegend,
    legendPos: "t",
    legendFontFace: p.font,
    legendFontSize: fontSize,
    legendColor: hex(p.text),
    showValue: obj.showValues && !radial,
    showPercent: radial && obj.showValues,
    showLabel: false,
    dataLabelColor: radial ? "FFFFFF" : hex(p.muted),
    dataLabelFontFace: p.font,
    dataLabelFontSize: fontSize,
    dataLabelFormatCode: radial ? "0%" : numberFormat,
    dataLabelPosition: radial ? "ctr" : obj.chartType === "line" || obj.chartType === "area" ? "t" : "outEnd",
  };
  if (!radial) {
    Object.assign(options, {
      barDir: obj.chartType === "bar" ? "bar" : "col",
      barGrouping: obj.stacked ? "stacked" : "clustered",
      barGapWidthPct: 40,
      catAxisLabelColor: hex(p.text),
      catAxisLabelFontFace: p.font,
      catAxisLabelFontSize: fontSize,
      catAxisOrientation: obj.chartType === "bar" ? "maxMin" : "minMax",
      valAxisLabelColor: hex(p.muted),
      valAxisLabelFontFace: p.font,
      valAxisLabelFontSize: fontSize,
      valAxisLabelFormatCode: numberFormat,
      valGridLine: { color: hex(p.line), size: 1 },
      catGridLine: { style: "none" },
      lineSize: 2,
      lineDataSymbol: obj.chartType === "line" ? "circle" : "none",
    });
    if (obj.chartType === "area" && obj.stacked) options.barGrouping = "stacked";
  } else if (obj.chartType === "doughnut") options.holeSize = 55;
  return { type, data, options };
}
