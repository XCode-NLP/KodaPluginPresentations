import assert from "node:assert/strict";
import { test } from "vitest";
import { fields, initial, objectAt } from "./helpers";
import PptxGenJS from "pptxgenjs";
import { apply } from "../src/model";
import type { ChartObject, ChartType } from "../src/types";
import { chartSvg, pptxChart } from "../src/chart";
import { buildLayoutSlide } from "../src/design/layouts";
import { resolvedTheme } from "../src/design/themes";

const revenue = {
  kind: "chart" as const,
  chartType: "column" as ChartType,
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
  const chart = objectAt(next, -1);
  assert.equal(chart.kind, "chart");
  assert.equal(chart.showLegend, true);
  assert.equal(chart.showValues, false);
  assert.equal(chart.fill, "transparent");
});

test("chart data is validated and editable through updateObject", () => {
  const state = initial();
  const slideId = state.document.slides[0].id;
  const added = apply(state, 1, [{ type: "addObject", slideId, object: revenue }]);
  const objectId = objectAt(added, -1).id;
  assert.throws(
    () => apply(added, 2, [{ type: "updateObject", slideId, objectId, patch: { categories: ["Q1", "Q2"] } }]),
    /Invalid chart/,
  );
  const line = apply(added, 2, [{ type: "updateObject", slideId, objectId, patch: { chartType: "line", unit: "%" } }]);
  assert.equal(objectAt(line, -1).chartType, "line");
  assert.throws(
    () => apply(added, 2, [{ type: "updateObject", slideId, objectId: objectAt(added, 0).id, patch: { chartType: "pie" } }]),
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
  for (const chartType of ["column", "bar", "line", "area", "pie", "doughnut"] as const) {
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
          id: "chart",
          fill: "transparent",
          stroke: "transparent",
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
  const obj: ChartObject = {
    ...revenue,
    id: "chart",
    fill: "transparent",
    stroke: "transparent",
    series: [revenue.series[0]],
    stacked: false,
    showLegend: false,
    showValues: true,
    unit: "%",
    highlight: 1,
  };
  const chart = pptxChart(pptx, obj, theme, { x: 1, y: 1, w: 6, h: 3.5 });
  assert.equal(chart.type, "bar");
  assert.deepEqual(chart.options.chartColors, [theme.colors.line, theme.colors.accent, theme.colors.line].map((c) => c.slice(1)));
  pptx.addSlide().addChart(chart.type, chart.data, chart.options);
  const zip = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  assert.ok(Buffer.from(zip).includes("ppt/charts/chart"));
});

test("the chart layout builds a native chart from one or several series", () => {
  const doc = initial().document;
  const theme = resolvedTheme("clean");
  doc.theme = theme;
  const multi = buildLayoutSlide(doc, theme, "chart", {
    title: "Выручка растёт",
    type: "line",
    categories: ["Q1", "Q2", "Q3"],
    series: revenue.series,
  });
  const chart = fields(multi.objects.find((obj) => obj.kind === "chart"));
  assert.equal(chart.chartType, "line");
  assert.equal(chart.series?.length, 2);
  const single = buildLayoutSlide(doc, theme, "chart", { title: "Отток", data: [{ label: "Н1", value: 40 }, { label: "Н2", value: 12 }] });
  assert.equal(fields(single.objects.find((obj) => obj.kind === "chart")).highlight, 0);
});
