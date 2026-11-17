import runner from "./runner.cjs";
import { apply, create, id, isId, object, selectionFor, travel, validate } from "./model.js";
import { DESIGN_GUIDE } from "./design/guide.js";
import { buildLayoutSlide, imageSlots, layoutCatalog } from "./design/layouts.js";
import { lintDeck } from "./design/lint.js";
import { catalog, themeOf } from "./design/themes.js";

const labels = {
  en: { title: "Title", subtitle: "Subtitle", text: "Text", caption: "Caption" },
  ru: { title: "Заголовок", subtitle: "Подзаголовок", text: "Текст", caption: "Подпись" },
};

function serialQueue() {
  let pending = Promise.resolve();
  return (fn) =>
    (...args) => {
      const result = pending.then(() => fn(...args));
      pending = result.catch(() => {});
      return result;
    };
}

export default {
  activate(ctx) {
    const serialize = serialQueue();

    const openJob = async (sessionId, job) => {
      const { projectPath } = await ctx.chat.getSession(sessionId);
      const name = `presentations-${Date.now()}-${Math.random().toString(36).slice(2)}`;
      const scratch = await ctx.sandboxes.create(name);
      const close = () =>
        ctx.sandboxes.remove(name).catch((error) => ctx.log.warn("Failed to remove job folder", String(error)));
      try {
        await ctx.projects.writeFile(scratch, "handler.cjs", runner, { expectedHash: null });
      } catch (error) {
        await close();
        throw error;
      }
      return { name, scratch, close, request: { ...job, root: projectPath } };
    };

    const runJob = async ({ scratch, request }) => {
      await ctx.projects.writeFile(scratch, "job.json", JSON.stringify(request), { expectedHash: null });
      const result = await ctx.process.runFile(scratch, "handler.cjs");
      if (result.timedOut) throw new Error("Presentation job exceeded 10 seconds");
      if (result.truncated) throw new Error("Presentation job output exceeded 64 KB");
      if (result.exitCode !== 0) throw new Error(String(result.output ?? "").trim() || `Presentation job failed: ${JSON.stringify(result).slice(0, 300)}`);
      const reply = JSON.parse(result.output);
      return reply.resultFile
        ? JSON.parse(await ctx.projects.readFile(scratch, "result.json"))
        : { ...reply, durationMs: result.durationMs };
    };

    const withJob = async (sessionId, request) => {
      const job = await openJob(sessionId, request);
      try {
        return await runJob(job);
      } finally {
        await job.close();
      }
    };

    const bindingKey = (sessionId) => `chat:${sessionId}`;
    const load = async (sessionId, documentId) => {
      if (!isId(documentId)) throw new Error("Invalid document ID");
      const state = await withJob(sessionId, { type: "load", documentId });
      if (
        !Number.isSafeInteger(state.version) ||
        state.version < 1 ||
        !Array.isArray(state.undo) ||
        !Array.isArray(state.redo) ||
        state.undo.length > 50 ||
        state.redo.length > 50
      )
        throw new Error("Invalid history");
      for (const doc of [state.document, ...state.undo, ...state.redo]) {
        if (validate(doc).id !== documentId) throw new Error("Document ID mismatch");
      }
      return state;
    };

    const describe = async (sessionId, state, selection) => {
      const { document, version, undo, redo } = state;
      selection = selectionFor(document, selection);
      const reference = `[presentation:${document.id}/slide:${selection.slideId}${selection.objectIds.map((key) => `/object:${key}`).join("")}]`;
      await ctx.chat.setContext(
        sessionId,
        `Active presentation selection: ${reference}. Read get_presentation and get_slide for current data before changing it. Explicit references in the prompt identify their original slide and objects. Before building or restyling slides, call get_design_system once and follow its workflow: plan the storyline, pick a theme, build slides with add_layout_slide, then check_presentation and render_slides to look at the result and fix it. Use apply_presentation_operations with the version returned by the read tools and theme role names instead of hex colors. To insert a picture, save it as a PNG in the project and call add_image with its path.`,
      );
      return { document, version, selection, reference, canUndo: !!undo.length, canRedo: !!redo.length };
    };

    const current = async (sessionId, wantedId) => {
      const binding = await ctx.storage.get(bindingKey(sessionId));
      if (!binding) return null;
      if (wantedId && binding.documentId !== wantedId) throw new Error("The active presentation changed");
      return describe(sessionId, await load(sessionId, binding.documentId), binding.selection);
    };

    const bind = async (sessionId, state, selection) => {
      selection = selectionFor(state.document, selection);
      const previous = await ctx.storage.get(bindingKey(sessionId));
      await ctx.storage.set(bindingKey(sessionId), { documentId: state.document.id, selection });
      if (previous?.documentId !== state.document.id) {
        await ctx.ui
          .post("binding", { sessionId, documentId: state.document.id, version: state.version })
          .catch((error) => ctx.log.warn("Could not notify the presentation panel", String(error)));
      }
      return describe(sessionId, state, selection);
    };

    const persist = async (sessionId, state, expectedVersion) => {
      while (JSON.stringify(state).length > 490_000) {
        const older = state.undo.length >= state.redo.length ? state.undo : state.redo;
        if (!older.length) throw new Error("Presentation exceeds the editor storage limit");
        older.shift();
      }
      await withJob(sessionId, { type: "save", state, expectedVersion, documentId: state.document.id });
      await ctx.ui
        .post("document", { documentId: state.document.id, version: state.version })
        .catch((error) => ctx.log.warn("Could not notify the presentation panel", String(error)));
    };

    const createDocument = async (sessionId, document) => {
      const state = { version: 1, document: validate(document), undo: [], redo: [] };
      await persist(sessionId, state, null);
      return bind(sessionId, state);
    };

    const change = async ({ sessionId, documentId, expectedVersion, operations, direction }) => {
      const binding = await ctx.storage.get(bindingKey(sessionId));
      if (binding?.documentId !== documentId) throw new Error("The active presentation changed");
      const state = await load(sessionId, documentId);
      const next = direction ? travel(state, expectedVersion, direction) : apply(state, expectedVersion, operations);
      await persist(sessionId, next, expectedVersion);
      return describe(sessionId, next, binding.selection);
    };

    const uploads = new Map();
    const upload = (sessionId, token) => {
      const entry = uploads.get(token);
      if (entry?.sessionId !== sessionId || !entry) throw new Error("Unknown upload");
      return entry;
    };
    const closeUpload = async (entry) => {
      uploads.delete(entry.job.name);
      clearTimeout(entry.timer);
      await entry.job.close();
    };
    const serializeUpload = serialQueue();
    const startUpload = serializeUpload(async ({ sessionId, documentId, file, resource, preview, count }) => {
      if (uploads.size >= 4 || !isId(documentId) || !Number.isInteger(count) || count < 1 || count > 342) {
        throw new Error("Invalid or oversized upload");
      }
      if (
        preview
          ? !isId(preview) || count > 57
          : resource
            ? !isId(resource)
            : !["presentation.pptx", "presentation.pdf"].includes(file)
      ) {
        throw new Error("Invalid presentation output");
      }
      const job = await openJob(sessionId, {
        type: preview ? "preview-write" : resource ? "asset-write" : "write",
        documentId,
        file,
        resource,
        preview,
        parts: count,
      });
      const entry = { sessionId, job, next: 0, count };
      entry.timer = setTimeout(() => void expireUpload(entry), 120_000);
      uploads.set(job.name, entry);
      return job.name;
    });
    const expireUpload = serializeUpload(async (entry) => {
      if (uploads.has(entry.job.name)) await closeUpload(entry);
    });
    ctx.ui.handle("readChunk", ({ sessionId, source, offset, stamp }) =>
      withJob(sessionId, { type: "read", source, offset, stamp }),
    );
    ctx.ui.handle("startUpload", startUpload);
    ctx.ui.handle(
      "appendUpload",
      serializeUpload(async ({ sessionId, token, index, part }) => {
        const entry = upload(sessionId, token);
        if (
          index !== entry.next ||
          index >= entry.count ||
          typeof part !== "string" ||
          part.length > 48 * 1024 ||
          part.length % 4 !== 0 ||
          !/^[a-z0-9+/]+={0,2}$/i.test(part)
        )
          throw new Error("Invalid upload chunk");
        await ctx.projects.writeFile(entry.job.scratch, `part-${index}.txt`, part, { expectedHash: null });
        entry.next++;
      }),
    );
    ctx.ui.handle(
      "finishUpload",
      serializeUpload(async ({ sessionId, token }) => {
        const entry = upload(sessionId, token);
        try {
          if (entry.next !== entry.count) throw new Error("Incomplete upload");
          return await runJob(entry.job);
        } finally {
          await closeUpload(entry);
        }
      }),
    );
    ctx.ui.handle(
      "abortUpload",
      serializeUpload(async ({ sessionId, token }) => {
        if (uploads.has(token)) await closeUpload(upload(sessionId, token));
      }),
    );
    ctx.ui.handle("readAsset", ({ sessionId, documentId, resource }) =>
      withJob(sessionId, { type: "asset-read", documentId, resource }),
    );
    ctx.ui.handle(
      "state",
      serialize(({ sessionId }) => current(sessionId)),
    );
    ctx.ui.handle("list", async ({ sessionId }) => (await withJob(sessionId, { type: "list" })).documents);
    ctx.ui.handle(
      "open",
      serialize(async ({ sessionId, documentId }) => bind(sessionId, await load(sessionId, documentId))),
    );
    ctx.ui.handle(
      "create",
      serialize(({ sessionId, title, language }) =>
        createDocument(sessionId, create(title, labels[language] ?? labels.en)),
      ),
    );
    ctx.ui.handle(
      "import",
      serialize(({ sessionId, document }) => createDocument(sessionId, document)),
    );
    ctx.ui.handle("change", serialize(change));
    ctx.ui.handle(
      "select",
      serialize(async ({ sessionId, documentId, selection }) => {
        const state = await load(sessionId, documentId);
        const binding = await ctx.storage.get(bindingKey(sessionId));
        if (binding?.documentId !== documentId) throw new Error("The active presentation changed");
        return bind(sessionId, state, selection);
      }),
    );
    ctx.ui.handle(
      "discuss",
      serialize(async ({ sessionId, documentId }) => {
        const result = await current(sessionId, documentId);
        if (!result) throw new Error("No active presentation");
        await ctx.chat.insertIntoComposer(result.reference, sessionId);
      }),
    );
    ctx.ui.handle("reveal", async ({ sessionId }) => {
      const { projectPath } = await ctx.chat.getSession(sessionId);
      await ctx.projects.reveal(projectPath);
    });

    const result = (value) => ({ content: JSON.stringify(value), structured: value });
    const touched = (operations) => [...new Set((operations ?? []).map((op) => op.slideId ?? op.slide?.id).filter(Boolean))];
    // Lint the full document; the summary drops object geometry.
    const withIssues = (value, slideIds) => {
      const known = slideIds.filter((key) => value.document.slides.some(({ id: slideId }) => slideId === key));
      return known.length ? { ...summary(value), issues: lintDeck(value.document, known) } : summary(value);
    };
    const summary = ({ document, ...info }) => ({
      ...info,
      document: {
        ...document,
        theme: { id: themeOf(document).id, fonts: themeOf(document).fonts },
        slides: document.slides.map(({ id, objects, warnings, layout, notes }) => ({
          id,
          ...(layout && { layout }),
          ...(notes && { notes: notes.slice(0, 120) }),
          objects: objects.map(({ id, kind, text }) => ({ id, kind, ...(text && { preview: text.slice(0, 120) }) })),
          ...(warnings && { warnings }),
        })),
      },
    });
    ctx.tools.register(
      "create_presentation",
      serialize(async ({ title }, { sessionId }) =>
        result(summary(await createDocument(sessionId, create(title, labels[ctx.env.language] ?? labels.en)))),
      ),
    );
    ctx.tools.register(
      "get_presentation",
      serialize(async (_args, { sessionId }) => {
        const value = await current(sessionId);
        return result(value ? summary(value) : { document: null });
      }),
    );
    ctx.tools.register(
      "get_slide",
      serialize(async ({ documentId, slideId }, { sessionId }) => {
        const value = await current(sessionId, documentId);
        const slide = value?.document.slides.find(({ id }) => id === slideId);
        if (!slide) throw new Error("Slide not found");
        return result({ documentId, version: value.version, slide });
      }),
    );
    ctx.tools.register(
      "apply_presentation_operations",
      serialize(async (args, { sessionId }) =>
        result(withIssues(await change({ ...args, sessionId, direction: null }), touched(args.operations))),
      ),
    );
    ctx.tools.register(
      "add_image",
      serialize(async ({ documentId, slideId, expectedVersion, source, x, y, w, h, index }, { sessionId }) => {
        const value = await current(sessionId, documentId);
        if (!value) throw new Error("No active presentation");
        if (value.version !== expectedVersion) throw new Error("Version conflict: read the presentation again");
        const { document } = value;
        if (!document.slides.some(({ id }) => id === slideId)) throw new Error("Slide not found");
        const resource = id();
        const { width, height } = await withJob(sessionId, { type: "asset-import", documentId, resource, source });
        // Missing sizes keep the aspect ratio; with no size at all the image fits half the slide.
        const ratio = width / height;
        const wide = w ?? (h ? h * ratio : Math.min(width, document.width * 0.5, document.height * 0.5 * ratio));
        const tall = h ?? wide / ratio;
        const obj = object("image", {
          resource,
          w: Math.max(1, wide),
          h: Math.max(1, tall),
          x: x ?? (document.width - wide) / 2,
          y: y ?? (document.height - tall) / 2,
          fill: "transparent",
        });
        try {
          const next = await change({
            sessionId,
            documentId,
            expectedVersion,
            direction: null,
            operations: [
              { type: "addObject", slideId, object: obj },
              ...(index === undefined ? [] : [{ type: "moveObject", slideId, objectId: obj.id, index }]),
            ],
          });
          return result({ ...summary(next), objectId: obj.id, resource, pixelWidth: width, pixelHeight: height });
        } catch (error) {
          await withJob(sessionId, { type: "asset-remove", documentId, resource }).catch(() => {});
          throw error;
        }
      }),
    );
    ctx.tools.register(
      "undo_presentation_change",
      serialize(async (args, { sessionId }) =>
        result(summary(await change({ ...args, sessionId, direction: "undo" }))),
      ),
    );
    ctx.tools.register("get_design_system", async (_args, { sessionId }) => {
      const value = sessionId ? await serialize(() => current(sessionId))().catch(() => null) : null;
      return result({
        guide: DESIGN_GUIDE,
        currentTheme: value ? themeOf(value.document).id : null,
        themes: catalog(),
        layouts: layoutCatalog(),
      });
    });
    ctx.tools.register(
      "set_presentation_theme",
      serialize(async ({ documentId, expectedVersion, theme, colors, fonts }, { sessionId }) => {
        const next = await change({
          sessionId,
          documentId,
          expectedVersion,
          direction: null,
          operations: [{ type: "setTheme", theme, ...(colors && { colors }), ...(fonts && { fonts }) }],
        });
        return result({ ...summary(next), issues: lintDeck(next.document) });
      }),
    );

    const importImage = async (sessionId, documentId, source) => {
      const resource = id();
      const { width, height } = await withJob(sessionId, { type: "asset-import", documentId, resource, source });
      return { resource, width, height };
    };
    ctx.tools.register(
      "add_layout_slide",
      serialize(async ({ documentId, expectedVersion, layout, slots, tone, index, notes, replaceSlideId }, { sessionId }) => {
        const value = await current(sessionId, documentId);
        if (!value) throw new Error("No active presentation");
        if (value.version !== expectedVersion) throw new Error("Version conflict: read the presentation again");
        const { document } = value;
        const replaced = replaceSlideId ? document.slides.findIndex(({ id: key }) => key === replaceSlideId) : -1;
        if (replaceSlideId && replaced < 0) throw new Error("Slide to replace not found");
        const images = {};
        try {
          for (const [slot, source] of Object.entries(imageSlots(layout, slots))) {
            images[slot] = await importImage(sessionId, documentId, source);
          }
          const slide = buildLayoutSlide(document, themeOf(document), layout, slots, { images, tone });
          if (typeof notes === "string" && notes.trim()) slide.notes = notes.trim().slice(0, 4000);
          const position = replaced >= 0 ? replaced : index;
          const next = await change({
            sessionId,
            documentId,
            expectedVersion,
            direction: null,
            operations: [
              { type: "addSlide", slide, ...(position === undefined ? {} : { index: Math.min(position, document.slides.length) }) },
              ...(replaced >= 0 ? [{ type: "deleteSlide", slideId: replaceSlideId }] : []),
            ],
          });
          return result({ ...summary(next), slideId: slide.id, issues: lintDeck(next.document, [slide.id]) });
        } catch (error) {
          for (const { resource } of Object.values(images)) {
            await withJob(sessionId, { type: "asset-remove", documentId, resource }).catch(() => {});
          }
          throw error;
        }
      }),
    );
    ctx.tools.register(
      "check_presentation",
      serialize(async ({ documentId, slideIds }, { sessionId }) => {
        const value = await current(sessionId, documentId);
        if (!value) throw new Error("No active presentation");
        return result({ documentId, version: value.version, ...lintDeck(value.document, slideIds) });
      }),
    );

    // Rendering needs the DOM, so the open panel draws the slides and saves PNG files into the
    // project; the agent then looks at them with its own file reading tool.
    const renders = new Map();
    ctx.ui.handle("renderAck", ({ sessionId, requestId }) => {
      const entry = renders.get(requestId);
      if (entry?.sessionId === sessionId) entry.ack();
    });
    ctx.ui.handle("rendered", ({ sessionId, requestId, files, error }) => {
      const entry = renders.get(requestId);
      if (entry?.sessionId !== sessionId) return;
      if (error) entry.reject(new Error(String(error)));
      else entry.resolve(files);
    });
    ctx.tools.register("render_slides", async ({ documentId, slideIds }, { sessionId }) => {
      const value = await serialize(() => current(sessionId, documentId))();
      if (!value) throw new Error("No active presentation");
      const all = value.document.slides.map(({ id: key }) => key);
      const wanted = slideIds?.length ? slideIds.filter((key) => all.includes(key)) : all;
      if (!wanted.length) throw new Error("Slide not found");
      const chosen = wanted.slice(0, 8);
      const requestId = id();
      let ack;
      let settle;
      const acked = new Promise((resolve) => (ack = resolve));
      const done = new Promise((resolve, reject) => (settle = { resolve, reject }));
      renders.set(requestId, { sessionId, ack, ...settle });
      const timeout = (ms, message) => new Promise((_, reject) => setTimeout(() => reject(new Error(message)), ms));
      try {
        await ctx.ui.post("render", { sessionId, requestId, documentId, version: value.version, slideIds: chosen });
        await Promise.race([
          acked,
          timeout(5000, "Rendering needs the Presentations panel open in this chat. Ask the user to open it, and meanwhile rely on check_presentation."),
        ]);
        const files = await Promise.race([done, timeout(90_000, "Rendering timed out")]);
        return result({
          documentId,
          version: value.version,
          files,
          ...(wanted.length > chosen.length && { skipped: wanted.length - chosen.length }),
          next: "Read each PNG with your file reading tool, compare it with the design guide, fix what looks wrong and render again.",
        });
      } finally {
        renders.delete(requestId);
      }
    });
  },
};
