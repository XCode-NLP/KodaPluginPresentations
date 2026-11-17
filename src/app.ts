import { id, object, selectionFor, slideFor } from "./model";
import { readSource, sendBinary, snapshot } from "./media";
import { themeOf } from "./design/themes";
import { exportPdf, exportPowerPoint, importDeck } from "./office";
import { strings } from "./strings";
import { chartSvg } from "./chart";
import type {} from "@kodadev/desktop-plugin-sdk/ui";
import type { ObjectKind, PresentationDocument, Selection, Slide, SlideObject } from "./types";

/** Elements are looked up by id from ui/index.html; one type covers the inputs, selects and dialogs used. */
type Control = HTMLElement & HTMLInputElement & HTMLDialogElement;
interface ViewState {
  document: PresentationDocument;
  version: number;
  selection: Selection;
  reference: string;
  canUndo: boolean;
  canRedo: boolean;
}
interface ChangeParams {
  sessionId: string;
  documentId: string;
  expectedVersion: number;
  operations?: unknown[] | null;
  direction?: "undo" | "redo" | null;
}
interface Drag {
  base: ViewState;
  slideId: string;
  objectId: string;
  objects: SlideObject[];
  resize: boolean;
  startX: number;
  startY: number;
  scale: number;
  operations: unknown[];
  pointerId?: number;
}
interface RenderRequest {
  sessionId: string;
  requestId: string;
  documentId: string;
  slideIds: string[];
}
type Language = keyof typeof strings;
type StringKey = keyof (typeof strings)["en"];

const errorText = (error: unknown) => (error instanceof Error ? error.message : String(error));

const plugin = window.kodaPlugin;
if (!plugin) throw new Error("Open this plugin inside Koda Desktop.");
const bridge = plugin;
const $ = (id: string): Control => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element #${id}`);
  return element as Control;
};
let language: Language = "ru";
let sessionId: string | null = null;
let generation = 0;
let state: ViewState | null = null;
let busy = false;
let retry: ChangeParams | null = null;
let drag: Drag | null = null;
let refreshNeeded = false;
const assets = new Map<string, Promise<string>>();
const t = (key: string, detail?: string): string =>
  [strings[language][key as StringKey] ?? key, detail].filter(Boolean).join(" · ");
/** The open document; handlers that need it run only while the editor is shown. */
const view = (): ViewState => {
  if (!state) throw new Error("No active presentation");
  return state;
};
const currentSlide = (): Slide => {
  const { document: doc, selection } = view();
  return doc.slides.find((slide) => slide.id === selection.slideId) ?? doc.slides[0];
};
const selected = (): SlideObject[] => {
  if (!state) return [];
  const ids = state.selection.objectIds;
  return currentSlide().objects.filter((obj) => ids.includes(obj.id));
};
const context = () => ({ sessionId: sessionId ?? "", documentId: view().document.id });

function status(message: string, error = false) {
  $("status").textContent = error ? "" : message;
  $("status").toggleAttribute("data-busy", busy);
  $("status").toggleAttribute("data-error", error);
  $("retry").hidden = !retry;
  $("feedback-text").textContent = error ? message : "";
  $("feedback").toggleAttribute("data-error", error);
  syncFeedback();
}

function syncFeedback() {
  $("feedback").hidden = !$("feedback-text").textContent && $("retry").hidden && $("reveal").hidden;
}

function closeMenus(except?: string) {
  for (const [button, menu] of [
    ["file-menu-button", "file-menu"],
    ["export-button", "export-menu"],
    ["add-slide-button", "layout-menu"],
  ]) {
    if (menu === except) continue;
    $(menu).hidden = true;
    $(button).setAttribute("aria-expanded", "false");
  }
}

for (const [button, menu] of [
  ["file-menu-button", "file-menu"],
  ["export-button", "export-menu"],
  ["add-slide-button", "layout-menu"],
]) {
  $(button).onclick = (event: MouseEvent) => {
    event.stopPropagation();
    const open = $(menu).hidden;
    closeMenus(menu);
    $(menu).hidden = !open;
    $(button).setAttribute("aria-expanded", String(open));
    if (open) $(menu).querySelector<HTMLButtonElement>("button:enabled")?.focus();
  };
  $(menu).addEventListener("click", () => closeMenus());
}
document.addEventListener("click", () => closeMenus());
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeMenus();
});
for (const card of document.querySelectorAll<HTMLElement>("[data-proxy]")) card.onclick = () => $(card.dataset.proxy ?? "").click();
for (const button of document.querySelectorAll<HTMLElement>("[data-kind]"))
  button.onclick = () => {
    $("kind").value = button.dataset.kind ?? "";
    $("add-object").click();
  };

function localize() {
  document.documentElement.lang = language;
  document.title = t("heading");
  for (const element of document.querySelectorAll<HTMLElement>("[data-text]")) element.textContent = t(element.dataset.text ?? "");
  for (const element of document.querySelectorAll<HTMLElement>("[data-label]")) {
    element.title = t(element.dataset.label ?? "");
    element.setAttribute("aria-label", element.title);
  }
  const options = (element: Control, keys: string[], suffix = "") => {
    const previous = element.value;
    element.replaceChildren(...keys.map((key) => new Option(t(key + suffix), key)));
    if (keys.includes(previous)) element.value = previous;
  };
  options($("layout"), ["title", "text", "columns", "caption"], "Layout");
  $("layout-menu").replaceChildren(
    ...["title", "text", "columns", "caption"].map((key) => {
      const item = document.createElement("button");
      item.setAttribute("role", "menuitem");
      item.textContent = t(key + "Layout");
      item.onclick = () => {
        $("layout").value = key;
        $("add-slide").click();
      };
      return item;
    }),
  );
  options($("kind"), ["text", "rect", "ellipse", "line", "image"]);
  options($("align"), ["left", "center", "right"]);
}

async function refresh() {
  const epoch = generation;
  const next = sessionId ? ((await bridge.call("state", { sessionId })) as ViewState | null) : null;
  if (epoch === generation) {
    state = next;
    render();
  }
}

async function action(fn: (active: string) => Promise<unknown>): Promise<void> {
  if (busy) return;
  const epoch = generation;
  const active = sessionId;
  if (!active) return status(t("chat"), true);
  busy = true;
  $("editor").setAttribute("aria-busy", "true");
  status(t("working"));
  $("status").setAttribute("data-busy", "");
  try {
    const result = await fn(active);
    if (epoch === generation) status(typeof result === "string" ? result : state ? t("saved") : "");
  } catch (error) {
    if (epoch === generation) {
      const message = errorText(error);
      if (retry && /Version conflict|active presentation changed/i.test(message)) {
        retry = null;
        refreshNeeded = true;
      }
      status(retry ? `${t("unsaved")} · ${message}` : message, true);
    }
  } finally {
    busy = false;
    $("editor").removeAttribute("aria-busy");
    $("status").removeAttribute("data-busy");
    if (refreshNeeded && !drag && !sheet.querySelector('[contenteditable="plaintext-only"]')) {
      refreshNeeded = false;
      refresh().catch((error) => status(errorText(error), true));
    }
  }
}

async function accept(method: string, params: unknown) {
  const epoch = generation;
  const next = (await bridge.call(method, params)) as ViewState;
  if (epoch === generation) {
    state = next;
    render();
  }
}

function commit(operations: unknown[] | null, direction?: "undo" | "redo" | null, base: ViewState | null = state) {
  if (!base || !sessionId || (!operations?.length && !direction)) return;
  const params: ChangeParams = { sessionId, documentId: base.document.id, expectedVersion: base.version, operations, direction };
  return action(async () => {
    retry = params;
    await accept("change", params);
    if (retry === params) retry = null;
  });
}

async function select(slideId: string, objectIds: string[]) {
  if (!state) return;
  const selection = selectionFor(state.document, { slideId, objectIds });
  if (busy) return;
  const request = { ...context(), selection };
  const sameSlide = state.selection.slideId === selection.slideId;
  state = { ...state, selection };
  if (sameSlide) renderSelection();
  else render();
  try {
    await bridge.call("select", request);
  } catch (error) {
    if (sessionId === request.sessionId) status(errorText(error), true);
  }
}

function imageFor(resource: string, request: { sessionId: string; documentId: string } = context()): Promise<string> {
  const key = `${request.sessionId}/${request.documentId}/${resource}`;
  if (!assets.has(key)) {
    const promise = bridge
      .call("readAsset", { ...request, resource })
      .then((answer) => `data:image/png;base64,${(answer as { data: string }).data}`)
      .catch((error) => {
        assets.delete(key);
        throw error;
      });
    assets.set(key, promise);
    const oldest = assets.keys().next().value;
    if (assets.size > 24 && oldest !== undefined) assets.delete(oldest);
  }
  return assets.get(key)!;
}

function slideElement(
  doc: PresentationDocument,
  slide: Slide,
  editable = false,
  request: { sessionId: string; documentId: string } = context(),
): HTMLElement {
  const theme = themeOf(doc);
  const element = document.createElement("div");
  element.className = editable ? "slide editable" : "slide";
  Object.assign(element.style, {
    width: `${doc.width}px`,
    height: `${doc.height}px`,
    background: slide.background ?? theme.background,
  });
  for (const obj of slide.objects) {
    const node = document.createElement("div");
    node.className = "slide-object";
    node.dataset.object = obj.id;
    Object.assign(node.style, {
      left: `${obj.x}px`,
      top: `${obj.y}px`,
      width: `${obj.w}px`,
      height: `${obj.h}px`,
      background: obj.kind === "line" ? "transparent" : obj.fill,
      boxShadow: obj.kind !== "line" && obj.stroke !== "transparent" ? `inset 0 0 0 1px ${obj.stroke}` : "",
      borderRadius: obj.kind === "ellipse" ? "50%" : "",
    });
    if (obj.kind === "text") {
      const content = document.createElement("div");
      content.textContent = obj.text;
      Object.assign(content.style, {
        fontFamily: `"${obj.font.replace(/["\\]/g, "")}", sans-serif`,
        fontSize: `${obj.size}px`,
        color: obj.color,
        textAlign: obj.align,
        fontWeight: obj.bold ? "bold" : "normal",
        fontStyle: obj.italic ? "italic" : "normal",
      });
      node.append(content);
    } else if (obj.kind === "image") {
      const img = new Image();
      img.alt = "";
      img.dataset.resource = obj.resource;
      node.append(img);
      void imageFor(obj.resource, request)
        .then((url) => {
          img.src = url;
        })
        .catch(() => {});
    } else if (obj.kind === "chart") {
      // chartSvg escapes every label, so the markup carries no user-controlled tags.
      node.innerHTML = chartSvg(obj, theme);
    } else if (obj.kind === "line") {
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("width", "100%");
      svg.setAttribute("height", "100%");
      const line = document.createElementNS(svg.namespaceURI, "line");
      line.setAttribute("x1", "0");
      line.setAttribute("y1", "0");
      line.setAttribute("x2", "100%");
      line.setAttribute("y2", "100%");
      line.setAttribute("stroke", obj.stroke);
      line.setAttribute("stroke-width", "2");
      svg.append(line);
      node.append(svg);
    }
    if (editable) {
      node.tabIndex = 0;
      node.setAttribute("aria-label", (obj.kind === "text" && obj.text.slice(0, 80)) || t(obj.kind));
    }
    element.append(node);
  }
  return element;
}

function sizeSheet() {
  if (!state || !$("sheet").firstChild) return;
  const { width, height } = state.document;
  const workspace = $("workspace");
  const style = getComputedStyle(workspace);
  const availableWidth = workspace.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
  const availableHeight = workspace.clientHeight - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom);
  const scale = Math.max(0.05, Math.min(availableWidth / width, availableHeight / height, 1));
  $("sheet").style.width = `${width * scale}px`;
  $("sheet").style.height = `${height * scale}px`;
  ($("sheet").firstChild as HTMLElement).style.transform = `scale(${scale})`;
}

function render() {
  const doc = state?.document;
  $("empty").hidden = !!doc;
  $("editor").hidden = !doc;
  $("title").hidden = !doc;
  $("empty-message").textContent = t(sessionId ? "empty" : "chat");
  for (const name of ["undo", "redo", "export-anchor"]) $(name).hidden = !doc;
  for (const name of ["create", "open", "import"]) $(name).disabled = !sessionId;
  $("empty-actions").hidden = !sessionId;
  if (!state || !doc) return;
  if (document.activeElement !== $("title")) $("title").value = doc.title;
  $("undo").disabled = !state.canUndo;
  $("redo").disabled = !state.canRedo;
  const slide = currentSlide();
  const index = doc.slides.indexOf(slide);
  $("delete-slide").disabled = doc.slides.length < 2;
  $("back").disabled = index === 0;
  $("forward").disabled = index === doc.slides.length - 1;
  const filmstrip = $("filmstrip");
  const scroll = filmstrip.scrollTop;
  filmstrip.replaceChildren(
    ...doc.slides.map((item, i) => {
      const thumb = document.createElement("button");
      thumb.className = "thumb";
      thumb.setAttribute("aria-current", String(item.id === slide.id));
      thumb.setAttribute("aria-label", `${t("slide")} ${i + 1}`);
      thumb.onclick = () => void select(item.id, []);
      const frame = document.createElement("div");
      frame.className = "thumb-paper";
      const thumbWidth = filmstrip.clientWidth - 28 || 72;
      frame.style.width = `${thumbWidth}px`;
      frame.style.height = `${(thumbWidth * doc.height) / doc.width}px`;
      const miniature = slideElement(doc, item);
      miniature.style.transform = `scale(${thumbWidth / doc.width})`;
      frame.append(miniature);
      const folio = document.createElement("span");
      folio.textContent = String(i + 1).padStart(2, "0");
      thumb.append(frame, folio);
      return thumb;
    }),
  );
  filmstrip.scrollTop = scroll;
  filmstrip.querySelector('[aria-current="true"]')?.scrollIntoView({ block: "nearest" });
  $("position").textContent = `${t("slide")} ${index + 1} / ${doc.slides.length}`;
  $("sheet").replaceChildren(slideElement(doc, slide, true));
  sizeSheet();
  renderSelection();
  const report = doc.slides.flatMap((item, i) =>
    (item.warnings ?? []).map((note) => `${t("slide")} ${i + 1} · ${note}`),
  );
  $("warnings").hidden = !report.length;
  $("warnings")
    .querySelector("ul")
    ?.replaceChildren(
      ...report.map((note) => {
        const li = document.createElement("li");
        li.textContent = note;
        return li;
      }),
    );
}

function renderSelection() {
  if (!state) return;
  const { selection: current } = state;
  const slide = currentSlide();
  const index = state.document.slides.indexOf(slide);
  for (const node of $("sheet").querySelectorAll<HTMLElement>("[data-object]")) {
    const active = current.objectIds.includes(node.dataset.object ?? "");
    node.toggleAttribute("data-selected", active);
    if (active && !node.querySelector(".handle")) {
      const handle = document.createElement("span");
      handle.className = "handle";
      node.append(handle);
    }
    if (!active) node.querySelector(".handle")?.remove();
  }
  const selection = selected();
  $("remove-object").disabled = !selection.length;
  $("properties").hidden = !selection.length;
  $("selection").textContent = [
    `${t("slide")} ${index + 1}`,
    selection.length === 1
      ? (selection[0].kind === "text" && selection[0].text.slice(0, 45)) || t(selection[0].kind)
      : selection.length
        ? `${selection.length} ${t("selected")}`
        : "",
  ]
    .filter(Boolean)
    .join(" · ");
  const first = selection[0];
  if (first) {
    const fields = first as unknown as Record<string, unknown>;
    for (const name of ["font", "size", "color", "align", "bold", "italic", "fill", "stroke"]) {
      const input = $(name);
      const row = input.closest<HTMLElement>(".color-row") ?? input.closest<HTMLElement>("label");
      if (row) row.hidden = !selection.every((obj) => name in obj);
      if (input.type === "checkbox") input.checked = !!fields[name];
      else if (input.type === "number") input.value = String(Math.round(Number(fields[name]) * 10) / 10);
      else if (name in first) input.value = fields[name] === "transparent" ? "#000000" : String(fields[name]);
    }
    $("no-fill").checked = first.fill === "transparent";
    $("no-stroke").checked = first.stroke === "transparent";
    $("lower").disabled = selection.length !== 1 || slide.objects.indexOf(first) === 0;
    $("raise").disabled = selection.length !== 1 || slide.objects.indexOf(first) === slide.objects.length - 1;
  }
  $("edit-text").hidden = selection.length !== 1 || first.kind !== "text";
  for (const [group, key] of [
    ['[data-group="text"]', "font"],
    [".section:has(#fill)", "fill"],
  ]) {
    const element = document.querySelector<HTMLElement>(group);
    if (element) element.hidden = !selection.length || !selection.every((obj) => key in obj);
  }
  requestAnimationFrame(sizeSheet);
}

const update = (patch: Record<string, unknown>) =>
  commit(selected().map((obj) => ({ type: "updateObject", slideId: currentSlide().id, objectId: obj.id, patch })));

$("create").onclick = () =>
  action(async (active) => {
    await accept("create", { sessionId: active, title: t("untitled"), language });
    retry = null;
  });
$("open").onclick = () =>
  action(async (active) => {
    const documents = (await bridge.call("list", { sessionId: active })) as { id: string; title: string }[];
    if (active !== sessionId) return;
    if (!documents.length) return t("noDocs");
    $("documents").replaceChildren(...documents.map((doc) => new Option(doc.title, doc.id)));
    $("open-dialog").showModal();
  });
$("open-dialog")
  .querySelector("form")
  ?.addEventListener("submit", (event) => {
    event.preventDefault();
    $("open-dialog").close();
    if ((event.submitter as HTMLButtonElement | null)?.value === "ok")
      void action(async (active) => {
        await accept("open", { sessionId: active, documentId: $("documents").value });
        retry = null;
      });
  });
$("import").onclick = () => {
  if (!busy) $("import-dialog").showModal();
};
$("import-dialog")
  .querySelector("form")
  ?.addEventListener("submit", (event) => {
    event.preventDefault();
    $("import-dialog").close();
    if ((event.submitter as HTMLButtonElement | null)?.value !== "ok") return;
    void action(async (active) => {
      const epoch = generation;
      const source = $("source").value;
      const bytes = await readSource(active, source, bridge);
      const doc = await importDeck(
        bytes,
        source
          .split("/")
          .at(-1)!
          .replace(/\.pptx$/i, "")
          .slice(0, 200),
        $("stage"),
        async (documentId, resource, blob) => {
          if (epoch !== generation) throw new Error(t("chat"));
          await sendBinary(bridge, { sessionId: active, documentId, resource }, blob);
        },
        t,
      );
      if (epoch !== generation) return;
      await accept("import", { sessionId: active, document: doc });
      retry = null;
    });
  });
$("title").onchange = () => commit([{ type: "rename", title: $("title").value }]);
$("undo").onclick = () => commit(null, "undo");
$("redo").onclick = () => commit(null, "redo");
$("retry").onclick = () =>
  action(async () => {
    await accept("change", retry);
    retry = null;
  });
$("discuss").onclick = () => action(() => bridge.call("discuss", context()));
$("reveal").onclick = () => action((active) => bridge.call("reveal", { sessionId: active }));
$("add-slide").onclick = () => {
  const { width, height } = view().document;
  return commit([
    {
      type: "addSlide",
      slide: slideFor($("layout").value, width, height, {
        title: t("titleLabel"),
        subtitle: t("subtitle"),
        text: t("text"),
        caption: t("caption"),
      }, themeOf(view().document)),
    },
  ]);
};
$("duplicate").onclick = () => {
  const slide = currentSlide();
  const duplicate = { ...slide, id: id(), objects: slide.objects.map((obj) => ({ ...obj, id: id() })) };
  return commit([{ type: "addSlide", slide: duplicate, index: view().document.slides.indexOf(slide) + 1 }]);
};
$("delete-slide").onclick = () => commit([{ type: "deleteSlide", slideId: currentSlide().id }]);
for (const [name, offset] of [
  ["back", -1],
  ["forward", 1],
] as const) {
  $(name).onclick = () =>
    commit([
      { type: "moveSlide", slideId: currentSlide().id, index: view().document.slides.indexOf(currentSlide()) + offset },
    ]);
}
$("add-object").onclick = () => {
  const kind = $("kind").value as ObjectKind;
  if (kind === "image") return $("image-file").click();
  return commit([
    { type: "addObject", slideId: currentSlide().id, object: object(kind, kind === "text" ? { text: t("text") } : {}) },
  ]);
};
$("image-file").onchange = () => {
  const file = $("image-file").files?.[0];
  $("image-file").value = "";
  if (!file) return;
  void action(async () => {
    const request = context();
    const base = view();
    const slideId = currentSlide().id;
    if (!/^image\/(png|jpeg|webp)$/.test(file.type) || file.size > 8_000_000)
      throw new Error("Unsupported or oversized image");
    const bitmap = await createImageBitmap(file);
    const canvas = document.createElement("canvas");
    try {
      const factor = Math.min(1, 1200 / Math.max(bitmap.width, bitmap.height));
      canvas.width = Math.max(1, Math.round(bitmap.width * factor));
      canvas.height = Math.max(1, Math.round(bitmap.height * factor));
      canvas.getContext("2d")?.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    } finally {
      bitmap.close();
    }
    const resource = id();
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve));
    if (!blob) throw new Error("Could not encode image");
    await sendBinary(bridge, { ...request, resource }, blob);
    const w = Math.min(260, canvas.width);
    await accept("change", {
      ...request,
      expectedVersion: base.version,
      operations: [
        {
          type: "addObject",
          slideId,
          object: object("image", {
            resource,
            w,
            h: Math.max(1, (w * canvas.height) / canvas.width),
            fill: "transparent",
          }),
        },
      ],
    });
  });
};
$("remove-object").onclick = () =>
  commit(selected().map((obj) => ({ type: "deleteObject", slideId: currentSlide().id, objectId: obj.id })));
for (const name of ["font", "size", "color", "align", "bold", "italic", "fill", "stroke"]) {
  $(name).onchange = () => {
    const input = $(name);
    const value =
      input.type === "checkbox" ? input.checked : input.type === "number" ? input.valueAsNumber : input.value;
    if (input.reportValidity()) return update({ [name]: value });
  };
}
for (const name of ["fill", "stroke"]) {
  $(`no-${name}`).onchange = (event: Event) => update({ [name]: (event.target as HTMLInputElement).checked ? "transparent" : $(name).value });
}
for (const [name, offset] of [
  ["lower", -1],
  ["raise", 1],
] as const) {
  $(name).onclick = () => {
    const slide = currentSlide();
    const obj = selected()[0];
    return commit([
      { type: "moveObject", slideId: slide.id, objectId: obj.id, index: slide.objects.indexOf(obj) + offset },
    ]);
  };
}

const sheet = $("sheet");
const targetOf = (event: Event) => event.target as HTMLElement;
sheet.addEventListener("pointerdown", (event) => {
  if (busy || !state || event.button !== 0 || targetOf(event).closest<HTMLElement>("[contenteditable]")?.isContentEditable) return;
  const editing = sheet.querySelector<HTMLElement>('[contenteditable="plaintext-only"]');
  if (editing) {
    editing.blur();
    return;
  }
  const node = targetOf(event).closest<HTMLElement>("[data-object]");
  if (!node) {
    void select(currentSlide().id, []);
    return;
  }
  const objectId = node.dataset.object ?? "";
  const slide = currentSlide();
  const oldIds = state.selection.objectIds;
  const additive = event.shiftKey || event.metaKey || event.ctrlKey;
  const ids = additive
    ? oldIds.includes(objectId)
      ? oldIds.filter((key) => key !== objectId)
      : [...oldIds, objectId]
    : oldIds.includes(objectId)
      ? oldIds
      : [objectId];
  if (ids !== oldIds) void select(slide.id, ids);
  if (additive) return;
  const resize = targetOf(event).classList.contains("handle");
  drag = {
    base: state,
    slideId: slide.id,
    objectId,
    objects: selected(),
    resize,
    startX: event.clientX,
    startY: event.clientY,
    scale: sheet.clientWidth / state.document.width,
    operations: [],
  };
  drag.pointerId = event.pointerId;
});
window.addEventListener("pointermove", (event) => {
  if (!drag || drag.pointerId !== event.pointerId) return;
  const current = drag;
  const dx = (event.clientX - drag.startX) / drag.scale;
  const dy = (event.clientY - drag.startY) / drag.scale;
  if (Math.abs(dx) + Math.abs(dy) < 2 && !drag.operations.length) return;
  if (!sheet.hasPointerCapture(event.pointerId)) sheet.setPointerCapture(event.pointerId);
  current.operations = current.objects
    .filter((obj) => !current.resize || obj.id === current.objectId)
    .map((obj) => {
      const patch: Record<string, number> = current.resize
        ? { w: Math.max(1, Math.min(8000, obj.w + dx)), h: Math.max(1, Math.min(8000, obj.h + dy)) }
        : { x: Math.max(-8000, Math.min(8000, obj.x + dx)), y: Math.max(-8000, Math.min(8000, obj.y + dy)) };
      const node = sheet.querySelector<HTMLElement>(`[data-object="${obj.id}"]`);
      const styles: Record<string, "left" | "top" | "width" | "height"> = { x: "left", y: "top", w: "width", h: "height" };
      for (const [key, value] of Object.entries(patch)) if (node) node.style[styles[key]] = `${value}px`;
      return { type: "updateObject", slideId: current.slideId, objectId: obj.id, patch };
    });
});
window.addEventListener("pointerup", (event) => {
  if (!drag || drag.pointerId !== event.pointerId) return;
  const completed = drag;
  drag = null;
  if (sheet.hasPointerCapture(event.pointerId)) sheet.releasePointerCapture(event.pointerId);
  if (completed.operations.length) void commit(completed.operations, null, completed.base);
  else if (refreshNeeded) {
    refreshNeeded = false;
    refresh().catch((error) => status(errorText(error), true));
  }
});
window.addEventListener("pointercancel", (event) => {
  if (drag?.pointerId === event.pointerId) {
    drag = null;
    render();
  }
});
function editText(node: HTMLElement | null) {
  if (busy || !state || !node || (node.firstChild as HTMLElement | null)?.isContentEditable) return;
  const slide = currentSlide();
  const obj = slide.objects.find((item) => item.id === node.dataset.object);
  if (obj?.kind !== "text") return;
  const base = state;
  const editor = node.firstChild as HTMLElement;
  editor.contentEditable = "plaintext-only";
  editor.setAttribute("aria-label", t("editText"));
  const listeners = new AbortController();
  const finish = (save: boolean) => {
    const value = editor.innerText;
    listeners.abort();
    editor.contentEditable = "false";
    editor.blur();
    if (!save) editor.textContent = obj.text;
    if (save && value !== obj.text) {
      void commit([{ type: "updateObject", slideId: slide.id, objectId: obj.id, patch: { text: value } }], null, base);
    } else if (refreshNeeded) {
      refreshNeeded = false;
      refresh().catch((error) => status(errorText(error), true));
    }
  };
  editor.addEventListener("blur", () => finish(true), { signal: listeners.signal });
  editor.addEventListener(
    "keydown",
    (event) => {
      if (event.key === "Escape" || (event.key === "Enter" && (event.metaKey || event.ctrlKey))) {
        event.preventDefault();
        event.stopPropagation();
        finish(event.key === "Enter");
      }
    },
    { signal: listeners.signal },
  );
  editor.focus();
}
sheet.addEventListener("dblclick", (event) => editText(targetOf(event).closest<HTMLElement>("[data-object]")));
sheet.addEventListener("keydown", (event) => {
  const node = targetOf(event).closest<HTMLElement>("[data-object]");
  if (!node || targetOf(event).isContentEditable) return;
  if (event.key === " " || event.key === "Enter") {
    event.preventDefault();
    void select(currentSlide().id, [node.dataset.object ?? ""]);
    if (event.key === "Enter") editText(node);
  }
});
$("edit-text").onclick = () => editText(sheet.querySelector<HTMLElement>("[data-selected]"));

for (const format of ["pptx", "pdf"]) {
  $(format).onclick = () =>
    action(async () => {
      const request = context();
      const doc = view().document;
      const file = `presentation.${format}`;
      const blob =
        format === "pptx"
          ? await exportPowerPoint(doc, (resource) => imageFor(resource, request))
          : await exportPdf(
              doc,
              async (slide) => {
                const element = slideElement(doc, slide, false, request);
                for (const img of element.querySelectorAll("img"))
                  img.src = await imageFor(img.dataset.resource ?? "", request);
                return element;
              },
              $("stage"),
            );
      await sendBinary(bridge, { ...request, file }, blob);
      if (request.sessionId === sessionId) {
        $("reveal").hidden = false;
        syncFeedback();
      }
      return `${t("saved")} · presentations/${doc.id}/exports/${file}`;
    });
}

// The agent's render_slides tool: draw the requested slides and save PNG previews into the project.
async function renderForAgent({ sessionId: requested, requestId, documentId, slideIds }: RenderRequest) {
  await bridge.call("renderAck", { sessionId: requested, requestId });
  try {
    const current = (await bridge.call("state", { sessionId: requested })) as ViewState | null;
    const doc = current?.document;
    if (!doc || doc.id !== documentId) throw new Error("The active presentation changed");
    const request = { sessionId: requested, documentId };
    const files: { slideId: string; slide: number; path: string }[] = [];
    const scale = Math.min(1.5, 1440 / doc.width);
    for (const slideId of slideIds) {
      const slide = doc.slides.find(({ id: key }) => key === slideId);
      if (!slide) continue;
      const element = slideElement(doc, slide, false, request);
      $("stage").replaceChildren(element);
      try {
        await Promise.all(
          [...element.querySelectorAll("img")].map(async (img) => {
            if (!img.src) img.src = await imageFor(img.dataset.resource ?? "", request);
            await img.decode();
          }),
        );
        const canvas = await snapshot(element, doc.width, doc.height, false, scale);
        const blob = await new Promise<Blob>((resolve, reject) =>
          canvas.toBlob((value) => (value ? resolve(value) : reject(new Error("PNG encoding failed")))),
        );
        await sendBinary(bridge, { ...request, preview: slideId }, blob);
        files.push({ slideId, slide: doc.slides.indexOf(slide) + 1, path: `presentations/${documentId}/previews/${slideId}.png` });
      } finally {
        $("stage").replaceChildren();
      }
    }
    await bridge.call("rendered", { sessionId: requested, requestId, files });
  } catch (error) {
    await bridge.call("rendered", { sessionId: requested, requestId, error: errorText(error) }).catch(() => {});
  }
}

new ResizeObserver(sizeSheet).observe($("workspace"));
bridge.subscribe((event, data) => {
  const value = data as Record<string, unknown>;
  if (event === "render" && value.sessionId === sessionId) void renderForAgent(data as RenderRequest);
  if (event === "theme") document.documentElement.dataset.theme = String(data);
  if (event === "language") {
    language = data === "en" ? "en" : "ru";
    localize();
    render();
  }
  if (event === "session") {
    generation++;
    sessionId = typeof data === "string" ? data : null;
    state = null;
    retry = null;
    drag = null;
    assets.clear();
    $("reveal").hidden = true;
    closeMenus();
    for (const dialog of document.querySelectorAll("dialog[open]")) (dialog as HTMLDialogElement).close("cancel");
    status("");
    render();
    refresh().catch((error) => status(errorText(error), true));
  }
  if (event === "binding" && value.sessionId === sessionId) {
    if (busy || drag || sheet.querySelector('[contenteditable="plaintext-only"]')) refreshNeeded = true;
    else refresh().catch((error) => status(errorText(error), true));
  }
  if (
    event === "document" &&
    sessionId &&
    (!state || (value.documentId === state.document.id && Number(value.version) > state.version))
  ) {
    if (busy || drag || sheet.querySelector('[contenteditable="plaintext-only"]')) refreshNeeded = true;
    else refresh().catch((error) => status(errorText(error), true));
  }
});
const info = await bridge.info();
language = info.language;
sessionId = info.sessionId;
document.documentElement.dataset.theme = info.theme;
localize();
render();
await refresh().catch((error) => status(errorText(error), true));
