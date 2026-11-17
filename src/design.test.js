import assert from "node:assert/strict";
import { test } from "vitest";
import { apply, create, validate } from "./model.js";
import { buildLayoutSlide, LAYOUT_IDS, layoutCatalog } from "./design/layouts.js";
import { lintDeck } from "./design/lint.js";
import { THEMES, contrast, resolvedTheme, themeOf } from "./design/themes.js";
import { fitText } from "./design/text.js";

const labels = { title: "Заголовок", subtitle: "Подзаголовок" };
const initial = () => ({ version: 1, document: create("Проект", labels), undo: [], redo: [] });

const SAMPLE = {
  title: { kicker: "Q3", title: "Онбординг v2 снизил отток", subtitle: "Итоги квартала и план", footer: "Команда роста · ноябрь" },
  section: { number: "02", title: "Что изменилось", subtitle: "Три решения" },
  statement: { statement: "Пользователи уходят в первые три дня", support: "Почти весь отток приходится на неделю после регистрации" },
  agenda: { title: "План", items: ["Проблема", "Решение", "Результаты", "Дальше"] },
  bullets: { title: "Отток упал на 30%", points: ["Короче первый шаг", "Подсказки в продукте", "Письма по поведению"], note: "Данные за июль–сентябрь" },
  text: { title: "Контекст", body: "Мы переписали онбординг и измерили эффект на когорте новых пользователей." },
  "two-columns": { title: "Две гипотезы", columns: [{ title: "Сложно", text: "Слишком много шагов" }, { title: "Непонятно", text: "Нет ценности сразу" }] },
  cards: { title: "Три принципа", cards: [{ title: "Быстро", text: "Ценность за минуту" }, { title: "Ясно", text: "Один шаг за раз" }, { title: "Лично", text: "Под задачу" }] },
  "big-number": { value: "−30%", label: "отток за первую неделю", context: "Против прошлого квартала на сопоставимой когорте" },
  stats: { title: "Ключевые метрики", stats: [{ value: "−30%", label: "отток" }, { value: "+18%", label: "активация" }, { value: "4 мин", label: "до ценности" }] },
  comparison: { title: "До и после", left: { title: "Было", points: ["9 шагов", "Нет подсказок"] }, right: { title: "Стало", points: ["3 шага", "Подсказки"] } },
  timeline: { title: "Дорожная карта", steps: [{ label: "Q1", title: "Исследование" }, { label: "Q2", title: "Прототип" }, { label: "Q3", title: "Запуск" }] },
  process: { title: "Как работает", steps: [{ title: "Регистрация", text: "Один экран" }, { title: "Цель", text: "Выбор задачи" }, { title: "Результат", text: "Первый успех" }] },
  quote: { quote: "Наконец-то понятно, с чего начать", author: "Анна Петрова", role: "Продакт, Acme" },
  "image-split": { title: "Новый экран", body: ["Один шаг", "Явная цель"] },
  "image-full": { caption: "Новый первый экран" },
  chart: { title: "Отток по неделям", data: [{ label: "Н1", value: 40 }, { label: "Н2", value: 22 }, { label: "Н3", value: 12 }], unit: "%", insight: "Первая неделя решает" },
  table: { title: "Сравнение тарифов", columns: ["Тариф", "Цена", "Лимит"], rows: [["Старт", "0", "3"], ["Про", "10", "∞"]], highlight: 1 },
  team: { title: "Команда", people: [{ name: "Анна Петрова", role: "Продакт" }, { name: "Иван Сидоров", role: "Дизайн" }] },
  closing: { title: "Спасибо", subtitle: "Вопросы?", contacts: "growth@example.com" },
};

test("every layout builds a valid slide without layout errors in every theme", () => {
  assert.deepEqual(new Set(Object.keys(SAMPLE)), new Set(LAYOUT_IDS));
  for (const theme of THEMES) {
    const doc = initial().document;
    const resolved = resolvedTheme(theme.id);
    doc.theme = resolved;
    doc.slides = LAYOUT_IDS.map((layout) => buildLayoutSlide(doc, resolved, layout, SAMPLE[layout]));
    validate(doc);
    const errors = lintDeck(doc).issues.filter((issue) => issue.severity === "error");
    assert.deepEqual(errors, [], `${theme.id}: ${JSON.stringify(errors)}`);
  }
});

test("list slots keep the whole text of every point", () => {
  const doc = initial().document;
  const theme = resolvedTheme("clean");
  const bullets = buildLayoutSlide(doc, theme, "bullets", SAMPLE.bullets);
  assert.ok(bullets.objects.some((obj) => obj.text?.includes("Письма по поведению")));
  const table = buildLayoutSlide(doc, theme, "table", SAMPLE.table);
  assert.ok(table.objects.some((obj) => obj.text === "Лимит"));
});

test("catalog exposes slots for the agent", () => {
  for (const entry of layoutCatalog()) assert.ok(entry.use && Object.keys(entry.slots).length, entry.id);
});

test("themes keep readable text on their own surfaces", () => {
  for (const { id, colors } of THEMES) {
    assert.ok(contrast(colors.text, colors.background) >= 7, `${id} text`);
    assert.ok(contrast(colors.muted, colors.background) >= 4.5, `${id} muted`);
    assert.ok(contrast(colors.text, colors.surface) >= 4.5, `${id} surface`);
    assert.ok(contrast(colors.onAccent, colors.accent) >= 4.5, `${id} onAccent`);
    assert.ok(contrast(colors.accent, colors.background) >= 4.5, `${id} accent`);
    assert.ok(contrast(colors.accent, colors.surface) >= 4.5, `${id} accent on surface`);
  }
});

test("role names resolve, persist and follow a theme switch", () => {
  const state = initial();
  const slideId = state.document.slides[0].id;
  const next = apply(state, 1, [
    { type: "addObject", slideId, object: { kind: "text", text: "Акцент", x: 64, y: 400, w: 400, h: 60, color: "accent", size: "h3" } },
  ]);
  const added = next.document.slides[0].objects.at(-1);
  assert.equal(added.color, resolvedTheme("clean").colors.accent);
  assert.deepEqual(added.roles, { font: "body", size: "h3", color: "accent" });
  const dark = apply(next, 2, [{ type: "setTheme", theme: "midnight" }]);
  const repainted = dark.document.slides[0].objects.at(-1);
  assert.equal(repainted.color, resolvedTheme("midnight").colors.accent);
  assert.equal(repainted.font, "Arial");
  assert.equal(dark.document.slides[0].objects[0].font, "Verdana");
  assert.equal(themeOf(dark.document).id, "midnight");
  const literal = apply(dark, 3, [{ type: "updateObject", slideId, objectId: added.id, patch: { color: "#123456" } }]);
  assert.equal(literal.document.slides[0].objects.at(-1).roles.color, undefined);
});

test("legacy documents without tokens get roles inferred on switch", () => {
  const state = initial();
  const doc = state.document;
  doc.theme = { background: "#ffffff" };
  for (const obj of doc.slides[0].objects) {
    delete obj.roles;
    obj.color = "#1c1f2e";
  }
  validate(doc);
  const next = apply(state, 1, [{ type: "setTheme", theme: "midnight" }]);
  assert.equal(next.document.slides[0].objects[0].color, resolvedTheme("midnight").colors.text);
});

test("lint reports overflow, contrast and placeholders", () => {
  const state = initial();
  const slideId = state.document.slides[0].id;
  const next = apply(state, 1, [
    { type: "addObject", slideId, object: { kind: "text", text: "очень ".repeat(80), x: 64, y: 300, w: 200, h: 40 } },
    { type: "addObject", slideId, object: { kind: "text", text: "Бледно", x: 600, y: 60, w: 200, h: 40, color: "#eeeeee" } },
  ]);
  const rules = lintDeck(next.document).issues.map((issue) => issue.rule);
  for (const rule of ["overflow", "contrast", "placeholder"]) assert.ok(rules.includes(rule), rule);
});

test("text fitting shrinks long text and never goes below the floor", () => {
  const fit = fitText("слово ".repeat(60), 300, 100, 40, "Arial", 13);
  assert.ok(fit.size < 40 && fit.size >= 13);
});
