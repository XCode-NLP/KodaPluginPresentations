import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { beforeAll, test, type TestContext } from "vitest";
import { bundleRunner } from "../scripts/runnerBundle.mjs";
import { labels } from "./helpers";
import { create, object } from "../src/model";
import type { ImageObject } from "../src/types";

// The test runs the same CommonJS text the plugin writes into its scratch folder.
let runner = "";
beforeAll(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "koda-presentation-runner-"));
  runner = path.join(dir, "handler.cjs");
  fs.writeFileSync(runner, await bundleRunner());
  return () => fs.rmSync(dir, { recursive: true, force: true });
});

function fixture(t: TestContext) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "koda-presentation-test-"));
  t.onTestFinished(() => fs.rmSync(root, { recursive: true, force: true }));
  return {
    root,
    run(request: Record<string, unknown>) {
      const cwd = fs.mkdtempSync(path.join(root, "job-"));
      fs.writeFileSync(path.join(cwd, "job.json"), JSON.stringify({ ...request, root }));
      const result = JSON.parse(
        execFileSync(process.execPath, [runner], { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }),
      );
      return result.resultFile ? JSON.parse(fs.readFileSync(path.join(cwd, "result.json"), "utf8")) : result;
    },
  };
}

test("slides are separate files and a conflicting or failed commit leaves the saved document intact", (t) => {
  const { root, run } = fixture(t);
  const document = create("Draft", labels);
  const state = { version: 1, document, undo: [], redo: [] };
  const documentId = document.id;
  run({ type: "save", documentId, state, expectedVersion: null });
  const index = path.join(root, "presentations", documentId, "document.json");
  const data = fs.readFileSync(index, "utf8");
  assert.ok(fs.existsSync(path.join(path.dirname(index), "slides", JSON.parse(data).document.slides[0] + ".json")));
  assert.deepEqual(run({ type: "load", documentId }), state);
  assert.throws(() => run({ type: "save", documentId, state, expectedVersion: null }), /Version conflict/);
  const changed = structuredClone(state);
  changed.version++;
  changed.document.slides[0].objects.push(object<ImageObject>("image", { resource: "missing" }));
  assert.throws(() => run({ type: "save", documentId, state: changed, expectedVersion: 1 }), /image resource/);
  assert.equal(fs.readFileSync(index, "utf8"), data);
});

test("binary reads enforce extension, chunk size, stale stamps and confinement", (t) => {
  const { root, run } = fixture(t);
  const bytes = Buffer.alloc(55_000, 221);
  fs.writeFileSync(path.join(root, "source.pptx"), bytes);
  const first = run({ type: "read", source: "source.pptx", offset: 0 });
  assert.equal(first.size, bytes.length);
  assert.ok(first.data.length < 64 * 1024);
  const offset = Buffer.from(first.data, "base64").length;
  const second = run({ type: "read", source: "source.pptx", offset, stamp: first.stamp });
  assert.deepEqual(Buffer.concat([Buffer.from(first.data, "base64"), Buffer.from(second.data, "base64")]), bytes);
  fs.appendFileSync(path.join(root, "source.pptx"), "a");
  assert.throws(
    () => run({ type: "read", source: "source.pptx", offset, stamp: first.stamp }),
    /changed during import/,
  );
  assert.throws(() => run({ type: "read", source: "../source.pptx", offset: 0 }), /Invalid relative path/);
  if (process.platform !== "win32") {
    fs.symlinkSync("source.pptx", path.join(root, "linked.pptx"));
    assert.throws(() => run({ type: "read", source: "linked.pptx", offset: 0 }), /without symlinks/);
  }
});

test("asset-import copies a project PNG into the document resources and reports its size", (t) => {
  const { root, run } = fixture(t);
  const png = Buffer.alloc(33);
  Buffer.from("89504e470d0a1a0a0000000d49484452", "hex").copy(png);
  png.writeUInt32BE(640, 16);
  png.writeUInt32BE(360, 20);
  fs.mkdirSync(path.join(root, "img"));
  fs.writeFileSync(path.join(root, "img", "chart.png"), png);
  fs.writeFileSync(path.join(root, "img", "fake.png"), "not a png");
  const job = { type: "asset-import", documentId: "doc-1", resource: "res-1", source: "img/chart.png" };
  assert.deepEqual(run(job), { width: 640, height: 360 });
  const target = path.join(root, "presentations", "doc-1", "resources", "res-1.png");
  assert.ok(fs.readFileSync(target).equals(png));
  assert.throws(() => run(job), /already exists/);
  assert.throws(() => run({ ...job, resource: "res-2", source: "img/fake.png" }), /not a PNG/);
  assert.throws(() => run({ ...job, resource: "res-3", source: "../outside.png" }), /Invalid relative path/);
  assert.throws(() => run({ ...job, resource: "res-4", source: "img/chart.jpg" }), /Expected a \.png/);
  run({ type: "asset-remove", documentId: "doc-1", resource: "res-1" });
  assert.ok(!fs.existsSync(target));
});
