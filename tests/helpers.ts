import { create } from "../src/model";
import type { ChartObject, HistoryState, ImageObject, Labels, SlideObject, TextObject } from "../src/types";

export const labels: Labels = { title: "Заголовок", subtitle: "Подзаголовок", text: "Текст", caption: "Подпись" };
export const initial = (): HistoryState => ({ version: 1, document: create("Проект", labels), undo: [], redo: [] });

/** Tests read fields of objects whose kind they already asserted. */
export type Fields = SlideObject &
  Partial<Omit<TextObject, "kind">> &
  Partial<Omit<ChartObject, "kind">> &
  Partial<Omit<ImageObject, "kind">>;
export const fields = (obj: SlideObject | undefined): Fields => {
  if (!obj) throw new Error("Object not found");
  return obj as Fields;
};
export const objectAt = (state: HistoryState, index: number, slide = 0): Fields =>
  fields(state.document.slides[slide].objects.at(index));
