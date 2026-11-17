import assert from "node:assert/strict";
import { test } from "vitest";
import PptxGenJS from "pptxgenjs";
import { apply, create } from "./model.js";
import { chartSvg, pptxChart } from "./chart.js";
import { buildLayoutSlide } from "./design/layouts.js";
import { resolvedTheme } from "./design/themes.js";

const labels = { title: "Заголовок", subtitle: "Подзаголовок" };
const initial = () => ({ version: 1, document: create("Проект", labels), undo: [], redo: [] });
const revenue = {
  kind: "chart",
  chartType: "column",
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "2025", values: [10, 14, 9] },
    { name: "2026", values: [12, 18, 21] },
  ],
  x: 64,
  y: 140,
  w: 600,
  h: 340,
};

test("an agent adds a partial chart and gets theme defaults", () => {
  const state = initial();
  const slideId = state.document.slides[0].id;
  const next = apply(state, 1, [{ type: "addObject", slideId, object: revenue }]);
  const chart = next.document.slides[0].objects.at(-1);
  assert.equal(chart.kind, "chart");
  assert.equal(chart.showLegend, true);
  assert.equal(chart.showValues, false);
  assert.equal(chart.fill, "transparent");
});

test("chart data is validated and editable through updateObject", () => {
  const state = initial();
  const slideId = state.document.slides[0].id;
  const added = apply(state, 1, [{ type: "addObject", slideId, object: revenue }]);
  const objectId = added.document.slides[0].objects.at(-1).id;
  assert.throws(
    () => apply(added, 2, [{ type: "updateObject", slideId, objectId, patch: { categories: ["Q1", "Q2"] } }]),
    /Invalid chart/,
  );
  const line = apply(added, 2, [{ type: "updateObject", slideId, objectId, patch: { chartType: "line", unit: "%" } }]);
  assert.equal(line.document.slides[0].objects.at(-1).chartType, "line");
  assert.throws(
    () => apply(added, 2, [{ type: "updateObject", slideId, objectId: added.document.slides[0].objects[0].id, patch: { chartType: "pie" } }]),
    /Invalid object patch/,
  );
  assert.throws(
    () =>
      apply(state, 1, [
        { type: "addObject", slideId, object: { ...revenue, chartType: "pie", series: [{ name: "a", values: [1, -2, 3] }] } },
      ]),
    /Invalid chart/,
  );
});

test("every chart type renders as escaped SVG in theme colors", () => {
  const theme = resolvedTheme("clean");
  for (const chartType of ["column", "bar", "line", "area", "pie", "doughnut"]) {
    for (const stacked of [false, true]) {
      const svg = chartSvg(
        {
          ...revenue,
          chartType,
          stacked,
          categories: ["<b>x</b>", "Q2", "Q3"],
          showLegend: true,
          showValues: true,
          unit: "%",
          highlight: null,
        },
        theme,
      );
      assert.match(svg, /^<svg [^>]+>.*<\/svg>$/s);
      assert.ok(!svg.includes("<b>"));
      assert.ok(svg.includes(theme.colors.chart1));
      assert.ok(!svg.includes("NaN"));
    }
  }
});

test("PowerPoint export writes a native chart with the highlight as point colors", async () => {
  const theme = resolvedTheme("clean");
  const pptx = new PptxGenJS();
  const obj = { ...revenue, series: [revenue.series[0]], stacked: false, showLegend: false, showValues: true, unit: "%", highlight: 1 };
  const chart = pptxChart(pptx, obj, theme, { x: 1, y: 1, w: 6, h: 3.5 });
  assert.equal(chart.type, pptx.ChartType.bar);
  assert.deepEqual(chart.options.chartColors, [theme.colors.line, theme.colors.accent, theme.colors.line].map((c) => c.slice(1)));
  pptx.addSlide().addChart(chart.type, chart.data, chart.options);
  const zip = await pptx.write({ outputType: "nodebuffer" });
  assert.ok(Buffer.from(zip).includes("ppt/charts/chart"));
});

test("the chart layout builds a native chart from one or several series", () => {
  const doc = initial().document;
  doc.theme = resolvedTheme("clean");
  const multi = buildLayoutSlide(doc, doc.theme, "chart", {
    title: "Выручка растёт",
    type: "line",
    categories: ["Q1", "Q2", "Q3"],
    series: revenue.series,
  });
  const chart = multi.objects.find((obj) => obj.kind === "chart");
  assert.equal(chart.chartType, "line");
  assert.equal(chart.series.length, 2);
  const single = buildLayoutSlide(doc, doc.theme, "chart", { title: "Отток", data: [{ label: "Н1", value: 40 }, { label: "Н2", value: 12 }] });
  assert.equal(single.objects.find((obj) => obj.kind === "chart").highlight, 0);
});
