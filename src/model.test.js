import assert from "node:assert/strict";
import { test } from "vitest";
import { apply, create, selectionFor, travel } from "./model.js";

const labels = { title: "Заголовок", subtitle: "Подзаголовок" };
const initial = () => ({ version: 1, document: create("Проект", labels), undo: [], redo: [] });
const rename = (title) => ({ type: "rename", title });

test("a failed operation rolls back the entire batch", () => {
  const state = initial();
  assert.throws(() => apply(state, 1, [rename("Changed"), { type: "deleteSlide", slideId: "missing" }]), /Operation 1 \(deleteSlide, fields: type, slideId\): Unknown slideId "missing". Available slideIds: /);
  assert.equal(state.document.title, "Проект");
  assert.equal(state.undo.length, 0);
});

test("explicit IDs keep changes on the intended object and versions detect concurrent edits", () => {
  const state = initial();
  const [slide] = state.document.slides;
  const [title, subtitle] = slide.objects;
  const selection = selectionFor(state.document, { slideId: slide.id, objectIds: [subtitle.id] });
  const batch = [{ type: "updateObject", slideId: slide.id, objectId: title.id, patch: { text: "Кратко" } }];
  const next = apply(state, 1, batch);
  assert.equal(next.document.slides[0].objects[0].text, "Кратко");
  assert.equal(next.document.slides[0].objects[1].text, "Подзаголовок");
  assert.deepEqual(selectionFor(next.document, selection), selection);
  assert.throws(() => apply(next, 1, batch), /Version conflict/);
  assert.throws(() => travel(next, 1, "undo"), /Version conflict/);
});

test("undo and redo are versioned actions, with history bounded to 50 edits", () => {
  let state = initial();
  for (let i = 1; i <= 60; i++) state = apply(state, state.version, [rename(String(i))]);
  assert.equal(state.undo.length, 50);
  const undone = travel(state, 61, "undo");
  assert.equal(undone.document.title, "59");
  assert.equal(undone.version, 62);
  const redone = travel(undone, 62, "redo");
  assert.equal(redone.document.title, "60");
  assert.equal(redone.version, 63);
  const edited = apply(undone, 62, [rename("New")]);
  assert.equal(edited.redo.length, 0);
  assert.throws(() => travel(edited, 63, "redo"), /Nothing to redo/);
});

test("operation errors name the field and the available IDs", () => {
  const state = initial();
  const slideId = state.document.slides[0].id;
  assert.throws(() => apply(state, 1, [{ type: "deleteSlide" }]), new RegExp(`Operation 0 \\(deleteSlide, fields: type\\): Missing slideId\\. Available slideIds: ${slideId}`));
  assert.throws(() => apply(state, 1, [{ type: "deleteSlide", id: slideId }]), /deleteSlide takes "slideId", not "id"/);
  assert.throws(() => apply(state, 1, [{ type: "removeSlide", slideId }]), /Unknown operation type "removeSlide"\. Expected one of: rename/);
  assert.throws(
    () => apply(state, 1, [{ type: "deleteObject", slideId, objectId: "nope" }]),
    new RegExp(`Unknown objectId "nope"\\. Available objectIds on slide ${slideId}:`),
  );
});

test("an object operation without slideId finds the slide by objectId", () => {
  const state = initial();
  const slide = state.document.slides[0];
  const object = slide.objects[0];
  const next = apply(state, 1, [{ type: "updateObject", objectId: object.id, patch: { text: "Новый" } }]);
  assert.equal(next.document.slides[0].objects[0].text, "Новый");
  assert.throws(() => apply(state, 1, [{ type: "updateObject", objectId: "nope", patch: {} }]), /Missing slideId/);
});
