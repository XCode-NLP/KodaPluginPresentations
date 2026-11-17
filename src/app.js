import { id, object, selectionFor, slideFor } from "./model.js";
import { readSource, sendBinary, snapshot } from "./media.js";
import { themeOf } from "./design/themes.js";
import { exportPdf, exportPowerPoint, importDeck } from "./office.js";
import { strings } from "./strings.js";
import { chartSvg } from "./chart.js";

const bridge = window.kodaPlugin;
const $ = (id) => document.getElementById(id);
let language = "ru";
let sessionId;
let generation = 0;
let state;
let busy = false;
let retry;
let drag;
let refreshNeeded = false;
const assets = new Map();
const t = (key, detail) => [strings[language][key], detail].filter(Boolean).join(" · ");
const currentSlide = () => state?.document.slides.find((slide) => slide.id === state.selection.slideId);
const selected = () => currentSlide()?.objects.filter((obj) => state.selection.objectIds.includes(obj.id)) ?? [];
const context = () => ({ sessionId, documentId: state.document.id });

function status(message, error = false) {
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

function closeMenus(except) {
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
  $(button).onclick = (event) => {
    event.stopPropagation();
    const open = $(menu).hidden;
    closeMenus(menu);
    $(menu).hidden = !open;
    $(button).setAttribute("aria-expanded", String(open));
    if (open) $(menu).querySelector("button:enabled")?.focus();
  };
  $(menu).addEventListener("click", () => closeMenus());
}
document.addEventListener("click", () => closeMenus());
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeMenus();
});
for (const card of document.querySelectorAll("[data-proxy]")) card.onclick = () => $(card.dataset.proxy).click();
for (const button of document.querySelectorAll("[data-kind]"))
  button.onclick = () => {
    $("kind").value = button.dataset.kind;
    $("add-object").click();
  };

function localize() {
  document.documentElement.lang = language;
  document.title = t("heading");
  for (const element of document.querySelectorAll("[data-text]")) element.textContent = t(element.dataset.text);
  for (const element of document.querySelectorAll("[data-label]")) {
    element.title = t(element.dataset.label);
    element.setAttribute("aria-label", element.title);
  }
  const options = (element, keys, suffix = "") => {
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
  const next = sessionId ? await bridge.call("state", { sessionId }) : null;
  if (epoch === generation) {
    state = next;
    render();
  }
}

async function action(fn) {
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
      if (retry && /Version conflict|active presentation changed/i.test(error.message)) {
        retry = null;
        refreshNeeded = true;
      }
      status(retry ? `${t("unsaved")} · ${error.message}` : error.message, true);
    }
  } finally {
    busy = false;
    $("editor").removeAttribute("aria-busy");
    $("status").removeAttribute("data-busy");
    if (refreshNeeded && !drag && !sheet.querySelector('[contenteditable="plaintext-only"]')) {
      refreshNeeded = false;
      refresh().catch((error) => status(error.message, true));
    }
  }
}

async function accept(method, params) {
  const epoch = generation;
  const next = await bridge.call(method, params);
  if (epoch === generation) {
    state = next;
    render();
  }
}

function commit(operations, direction, base = state) {
  if (!base || (!operations?.length && !direction)) return;
  const params = { sessionId, documentId: base.document.id, expectedVersion: base.version, operations, direction };
  return action(async () => {
    retry = params;
    await accept("change", params);
    if (retry === params) retry = null;
  });
}

async function select(slideId, objectIds) {
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
    if (sessionId === request.sessionId) status(error.message, true);
  }
}

function imageFor(resource, request = context()) {
  const key = `${request.sessionId}/${request.documentId}/${resource}`;
  if (!assets.has(key)) {
    const promise = bridge
      .call("readAsset", { ...request, resource })
      .then(({ data }) => `data:image/png;base64,${data}`)
      .catch((error) => {
        assets.delete(key);
        throw error;
      });
    assets.set(key, promise);
    if (assets.size > 24) assets.delete(assets.keys().next().value);
  }
  return assets.get(key);
}

function slideElement(doc, slide, editable = false, request = context()) {
  const element = document.createElement("div");
  element.className = editable ? "slide editable" : "slide";
  Object.assign(element.style, {
    width: `${doc.width}px`,
    height: `${doc.height}px`,
    background: slide.background ?? doc.theme.background,
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
      node.innerHTML = chartSvg(obj, doc.theme);
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
      node.setAttribute("aria-label", obj.text?.slice(0, 80) || t(obj.kind));
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
  $("sheet").firstChild.style.transform = `scale(${scale})`;
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
  if (!doc) return;
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
    .replaceChildren(
      ...report.map((note) => {
        const li = document.createElement("li");
        li.textContent = note;
        return li;
      }),
    );
}

function renderSelection() {
  const slide = currentSlide();
  const index = state.document.slides.indexOf(slide);
  for (const node of $("sheet").querySelectorAll("[data-object]")) {
    const active = state.selection.objectIds.includes(node.dataset.object);
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
      ? selection[0].text?.slice(0, 45) || t(selection[0].kind)
      : selection.length
        ? `${selection.length} ${t("selected")}`
        : "",
  ]
    .filter(Boolean)
    .join(" · ");
  const first = selection[0];
  if (first) {
    for (const name of ["font", "size", "color", "align", "bold", "italic", "fill", "stroke"]) {
      const input = $(name);
      (input.closest(".color-row") ?? input.closest("label")).hidden = !selection.every((obj) => name in obj);
      if (input.type === "checkbox") input.checked = !!first[name];
      else if (input.type === "number") input.value = String(Math.round(first[name] * 10) / 10);
      else if (name in first) input.value = first[name] === "transparent" ? "#000000" : first[name];
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
  ])
    document.querySelector(group).hidden = !selection.length || !selection.every((obj) => key in obj);
  requestAnimationFrame(sizeSheet);
}

const update = (patch) =>
  commit(selected().map((obj) => ({ type: "updateObject", slideId: currentSlide().id, objectId: obj.id, patch })));

$("create").onclick = () =>
  action(async (active) => {
    await accept("create", { sessionId: active, title: t("untitled"), language });
    retry = null;
  });
$("open").onclick = () =>
  action(async (active) => {
    const documents = await bridge.call("list", { sessionId: active });
    if (active !== sessionId) return;
    if (!documents.length) return t("noDocs");
    $("documents").replaceChildren(...documents.map((doc) => new Option(doc.title, doc.id)));
    $("open-dialog").showModal();
  });
$("open-dialog")
  .querySelector("form")
  .addEventListener("submit", (event) => {
    event.preventDefault();
    $("open-dialog").close();
    if (event.submitter?.value === "ok")
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
  .addEventListener("submit", (event) => {
    event.preventDefault();
    $("import-dialog").close();
    if (event.submitter?.value !== "ok") return;
    void action(async (active) => {
      const epoch = generation;
      const source = $("source").value;
      const bytes = await readSource(active, source, bridge);
      const doc = await importDeck(
        bytes,
        source
          .split("/")
          .at(-1)
          .replace(/\.pptx$/i, "")
          .slice(0, 200),
        $("stage"),
        (documentId, resource, blob) => {
          if (epoch !== generation) throw new Error(t("chat"));
          return sendBinary(bridge, { sessionId: active, documentId, resource }, blob);
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
  const { width, height } = state.document;
  return commit([
    {
      type: "addSlide",
      slide: slideFor($("layout").value, width, height, {
        title: t("titleLabel"),
        subtitle: t("subtitle"),
        text: t("text"),
        caption: t("caption"),
      }, themeOf(state.document)),
    },
  ]);
};
$("duplicate").onclick = () => {
  const slide = currentSlide();
  const duplicate = { ...slide, id: id(), objects: slide.objects.map((obj) => ({ ...obj, id: id() })) };
  return commit([{ type: "addSlide", slide: duplicate, index: state.document.slides.indexOf(slide) + 1 }]);
};
$("delete-slide").onclick = () => commit([{ type: "deleteSlide", slideId: currentSlide().id }]);
for (const [name, offset] of [
  ["back", -1],
  ["forward", 1],
]) {
  $(name).onclick = () =>
    commit([
      { type: "moveSlide", slideId: currentSlide().id, index: state.document.slides.indexOf(currentSlide()) + offset },
    ]);
}
$("add-object").onclick = () => {
  const kind = $("kind").value;
  if (kind === "image") return $("image-file").click();
  return commit([
    { type: "addObject", slideId: currentSlide().id, object: object(kind, kind === "text" ? { text: t("text") } : {}) },
  ]);
};
$("image-file").onchange = () => {
  const file = $("image-file").files[0];
  $("image-file").value = "";
  if (!file) return;
  void action(async () => {
    const request = context();
    const base = state;
    const slideId = currentSlide().id;
    if (!/^image\/(png|jpeg|webp)$/.test(file.type) || file.size > 8_000_000)
      throw new Error("Unsupported or oversized image");
    const bitmap = await createImageBitmap(file);
    const canvas = document.createElement("canvas");
    try {
      const factor = Math.min(1, 1200 / Math.max(bitmap.width, bitmap.height));
      canvas.width = Math.max(1, Math.round(bitmap.width * factor));
      canvas.height = Math.max(1, Math.round(bitmap.height * factor));
      canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    } finally {
      bitmap.close();
    }
    const resource = id();
    const blob = await new Promise((resolve) => canvas.toBlob(resolve));
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
  $(`no-${name}`).onchange = (event) => update({ [name]: event.target.checked ? "transparent" : $(name).value });
}
for (const [name, offset] of [
  ["lower", -1],
  ["raise", 1],
]) {
  $(name).onclick = () => {
    const slide = currentSlide();
    const obj = selected()[0];
    return commit([
      { type: "moveObject", slideId: slide.id, objectId: obj.id, index: slide.objects.indexOf(obj) + offset },
    ]);
  };
}

const sheet = $("sheet");
sheet.addEventListener("pointerdown", (event) => {
  if (busy || event.button !== 0 || event.target.closest("[contenteditable]")?.isContentEditable) return;
  const editing = sheet.querySelector('[contenteditable="plaintext-only"]');
  if (editing) {
    editing.blur();
    return;
  }
  const node = event.target.closest("[data-object]");
  if (!node) {
    void select(currentSlide().id, []);
    return;
  }
  const objectId = node.dataset.object;
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
  const resize = event.target.classList.contains("handle");
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
  const dx = (event.clientX - drag.startX) / drag.scale;
  const dy = (event.clientY - drag.startY) / drag.scale;
  if (Math.abs(dx) + Math.abs(dy) < 2 && !drag.operations.length) return;
  if (!sheet.hasPointerCapture(event.pointerId)) sheet.setPointerCapture(event.pointerId);
  drag.operations = drag.objects
    .filter((obj) => !drag.resize || obj.id === drag.objectId)
    .map((obj) => {
      const patch = drag.resize
        ? { w: Math.max(1, Math.min(8000, obj.w + dx)), h: Math.max(1, Math.min(8000, obj.h + dy)) }
        : { x: Math.max(-8000, Math.min(8000, obj.x + dx)), y: Math.max(-8000, Math.min(8000, obj.y + dy)) };
      const node = sheet.querySelector(`[data-object="${obj.id}"]`);
      for (const [key, value] of Object.entries(patch))
        node.style[{ x: "left", y: "top", w: "width", h: "height" }[key]] = `${value}px`;
      return { type: "updateObject", slideId: drag.slideId, objectId: obj.id, patch };
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
    refresh().catch((error) => status(error.message, true));
  }
});
window.addEventListener("pointercancel", (event) => {
  if (drag?.pointerId === event.pointerId) {
    drag = null;
    render();
  }
});
function editText(node) {
  if (busy || !node || node.firstChild?.isContentEditable) return;
  const slide = currentSlide();
  const obj = slide.objects.find((obj) => obj.id === node.dataset.object);
  if (obj?.kind !== "text") return;
  const base = state;
  const editor = node.firstChild;
  editor.contentEditable = "plaintext-only";
  editor.setAttribute("aria-label", t("editText"));
  const listeners = new AbortController();
  const finish = (save) => {
    const value = editor.innerText;
    listeners.abort();
    editor.contentEditable = "false";
    editor.blur();
    if (!save) editor.textContent = obj.text;
    if (save && value !== obj.text) {
      void commit([{ type: "updateObject", slideId: slide.id, objectId: obj.id, patch: { text: value } }], null, base);
    } else if (refreshNeeded) {
      refreshNeeded = false;
      refresh().catch((error) => status(error.message, true));
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
sheet.addEventListener("dblclick", (event) => editText(event.target.closest("[data-object]")));
sheet.addEventListener("keydown", (event) => {
  const node = event.target.closest("[data-object]");
  if (!node || event.target.isContentEditable) return;
  if (event.key === " " || event.key === "Enter") {
    event.preventDefault();
    void select(currentSlide().id, [node.dataset.object]);
    if (event.key === "Enter") editText(node);
  }
});
$("edit-text").onclick = () => editText(sheet.querySelector("[data-selected]"));

for (const format of ["pptx", "pdf"]) {
  $(format).onclick = () =>
    action(async () => {
      const request = context();
      const doc = state.document;
      const file = `presentation.${format}`;
      const blob =
        format === "pptx"
          ? await exportPowerPoint(doc, (resource) => imageFor(resource, request))
          : await exportPdf(
              doc,
              async (slide) => {
                const element = slideElement(doc, slide, false, request);
                for (const img of element.querySelectorAll("img"))
                  img.src = await imageFor(img.dataset.resource, request);
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
async function renderForAgent({ sessionId: requested, requestId, documentId, slideIds }) {
  await bridge.call("renderAck", { sessionId: requested, requestId });
  try {
    const { document: doc } = await bridge.call("state", { sessionId: requested });
    if (doc?.id !== documentId) throw new Error("The active presentation changed");
    const request = { sessionId: requested, documentId };
    const files = [];
    const scale = Math.min(1.5, 1440 / doc.width);
    for (const slideId of slideIds) {
      const slide = doc.slides.find(({ id: key }) => key === slideId);
      if (!slide) continue;
      const element = slideElement(doc, slide, false, request);
      $("stage").replaceChildren(element);
      try {
        await Promise.all(
          [...element.querySelectorAll("img")].map(async (img) => {
            if (!img.src) await imageFor(img.dataset.resource, request).then((url) => (img.src = url));
            await img.decode();
          }),
        );
        const canvas = await snapshot(element, doc.width, doc.height, false, scale);
        const blob = await new Promise((resolve, reject) =>
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
    await bridge.call("rendered", { sessionId: requested, requestId, error: error.message }).catch(() => {});
  }
}

new ResizeObserver(sizeSheet).observe($("workspace"));
bridge.subscribe((event, value) => {
  if (event === "render" && value.sessionId === sessionId) void renderForAgent(value);
  if (event === "theme") document.documentElement.dataset.theme = value;
  if (event === "language") {
    language = value;
    localize();
    render();
  }
  if (event === "session") {
    generation++;
    sessionId = value;
    state = retry = drag = null;
    assets.clear();
    $("reveal").hidden = true;
    closeMenus();
    for (const dialog of document.querySelectorAll("dialog[open]")) dialog.close("cancel");
    status("");
    render();
    refresh().catch((error) => status(error.message, true));
  }
  if (event === "binding" && value.sessionId === sessionId) {
    if (busy || drag || sheet.querySelector('[contenteditable="plaintext-only"]')) refreshNeeded = true;
    else refresh().catch((error) => status(error.message, true));
  }
  if (
    event === "document" &&
    sessionId &&
    (!state || (value.documentId === state.document.id && value.version > state.version))
  ) {
    if (busy || drag || sheet.querySelector('[contenteditable="plaintext-only"]')) refreshNeeded = true;
    else refresh().catch((error) => status(error.message, true));
  }
});
const info = await bridge.info();
language = info.language;
sessionId = info.sessionId;
document.documentElement.dataset.theme = info.theme;
localize();
render();
await refresh().catch((error) => status(error.message, true));
