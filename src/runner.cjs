const fs = require("node:fs");
const path = require("node:path");
const { createHash, randomUUID } = require("node:crypto");

const chunkSize = 45 * 1024;
const fileLimit = 12 * 1024 * 1024;
const isId = (value) => typeof value === "string" && /^[a-z0-9-]{1,64}$/.test(value);
const pngSignature = Buffer.from("89504e470d0a1a0a", "hex");

function inside(root, relative, createParents = false) {
  const parts = typeof relative === "string" ? relative.split("/") : [];
  if (
    !parts.length ||
    parts.some(
      (part) =>
        !part ||
        /[ .]$/.test(part) ||
        /[\\:<>"|?*\x00-\x1f]/.test(part) ||
        /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(part),
    )
  ) {
    throw new Error("Invalid relative path");
  }
  let current = root;
  for (const part of parts.slice(0, -1)) {
    current = path.join(current, part);
    if (createParents && !fs.existsSync(current)) {
      try {
        fs.mkdirSync(current);
      } catch (error) {
        if (error.code !== "EEXIST") throw error;
      }
    }
    if (!fs.lstatSync(current).isDirectory()) throw new Error("Expected a directory without symlinks");
  }
  const target = path.join(current, parts.at(-1));
  const stat = fs.lstatSync(target, { throwIfNoEntry: false });
  if (stat && !stat.isFile()) throw new Error("Expected a regular file without symlinks");
  return target;
}

function atomicWrite(target, bytes) {
  const temp = path.join(path.dirname(target), `.${randomUUID()}.tmp`);
  try {
    fs.writeFileSync(temp, bytes, { flag: "wx" });
    fs.renameSync(temp, target);
  } finally {
    fs.rmSync(temp, { force: true });
  }
}

function acquireLock(lock) {
  try {
    return fs.openSync(lock, "wx");
  } catch (error) {
    if (error.code !== "EEXIST") throw error;
    const age = Date.now() - fs.statSync(lock).mtimeMs;
    const pid = Number(fs.readFileSync(lock, "utf8"));
    // A killed handler can leave a lock behind. Let any active handler finish first.
    if (age < 30_000) throw new Error("The presentation is being saved. Try again.");
    if (Number.isSafeInteger(pid) && pid > 0) {
      try {
        process.kill(pid, 0);
      } catch (checkError) {
        if (checkError.code !== "ESRCH") throw checkError;
        fs.unlinkSync(lock);
        return fs.openSync(lock, "wx");
      }
      throw new Error("The presentation is being saved. Try again.");
    }
    fs.unlinkSync(lock);
    return fs.openSync(lock, "wx");
  }
}

function documentJob(root, job) {
  const base = `presentations/${job.documentId}`;
  const index = inside(root, `${base}/document.json`, job.type === "save");
  if (job.type === "load" && !fs.existsSync(index)) throw new Error("Presentation not found");
  if (job.type === "load") {
    for (let attempt = 0; attempt < 3; attempt++) {
      const serialized = fs.readFileSync(index, "utf8");
      const state = JSON.parse(serialized);
      const cache = new Map();
      const hydrate = (doc) => ({
        ...doc,
        slides: doc.slides.map((key) => {
          if (!/^[a-f0-9]{64}$/.test(key)) throw new Error("Invalid slide reference");
          if (!cache.has(key))
            cache.set(key, JSON.parse(fs.readFileSync(inside(root, `${base}/slides/${key}.json`), "utf8")));
          return cache.get(key);
        }),
      });
      try {
        const output = JSON.stringify({
          version: state.version,
          document: hydrate(state.document),
          undo: state.undo.map(hydrate),
          redo: state.redo.map(hydrate),
        });
        if (Buffer.byteLength(output) > 2_000_000) throw new Error("Presentation history exceeds storage limit");
        fs.writeFileSync("result.json", output, { flag: "wx" });
        return { resultFile: true };
      } catch (error) {
        if (error.code !== "ENOENT" || fs.readFileSync(index, "utf8") === serialized) throw error;
      }
    }
    throw new Error("Presentation changed during read. Try again.");
  }

  // An exclusive lock also protects against another Desktop process opening the same project.
  const lock = inside(root, `${base}/.write-lock`);
  const fd = acquireLock(lock);
  try {
    fs.writeFileSync(fd, String(process.pid));
    const prior = fs.existsSync(index) ? JSON.parse(fs.readFileSync(index, "utf8")) : null;
    if ((prior?.version ?? null) !== job.expectedVersion) throw new Error("Version conflict: reload the presentation");
    const next = job.state;
    const resources = new Set(
      next.document.slides.flatMap(({ objects }) =>
        objects.filter(({ kind }) => kind === "image").map(({ resource }) => resource),
      ),
    );
    for (const resource of resources) {
      if (!isId(resource)) throw new Error("Invalid image resource ID");
      try {
        if (!fs.existsSync(inside(root, `${base}/resources/${resource}.png`)))
          throw new Error("An image resource is missing");
      } catch (error) {
        if (error.code === "ENOENT") throw new Error("An image resource is missing");
        throw error;
      }
    }
    const keys = new Set();
    const pack = (doc) => ({
      ...doc,
      slides: doc.slides.map((slide) => {
        const bytes = JSON.stringify(slide);
        const key = createHash("sha256").update(bytes).digest("hex");
        if (!keys.has(key)) {
          const file = inside(root, `${base}/slides/${key}.json`, true);
          if (!fs.existsSync(file)) atomicWrite(file, bytes);
          keys.add(key);
        }
        return key;
      }),
    });
    const packed = {
      version: next.version,
      document: pack(next.document),
      undo: next.undo.map(pack),
      redo: next.redo.map(pack),
    };
    atomicWrite(index, JSON.stringify(packed));
    try {
      const folder = path.join(root, base, "slides");
      for (const file of fs.readdirSync(folder)) {
        if (/^[a-f0-9]{64}\.json$/.test(file) && !keys.has(file.slice(0, 64))) fs.rmSync(path.join(folder, file));
      }
    } catch {
      // A cleanup failure must not turn a committed save into an apparent failure.
    }
    return { saved: true };
  } finally {
    fs.closeSync(fd);
    fs.rmSync(lock, { force: true });
  }
}

function listDocuments(root) {
  const folder = path.join(root, "presentations");
  const stat = fs.lstatSync(folder, { throwIfNoEntry: false });
  if (stat && !stat.isDirectory()) throw new Error("The presentations folder must be a real directory");
  const documents = [];
  if (stat) {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true }).slice(0, 2000)) {
      if (!entry.isDirectory() || !isId(entry.name)) continue;
      try {
        const index = inside(root, `presentations/${entry.name}/document.json`);
        if (fs.statSync(index).size > 2_000_000) continue;
        const { document } = JSON.parse(fs.readFileSync(index, "utf8"));
        if (document.id === entry.name && typeof document.title === "string" && document.title.length <= 200) {
          documents.push({ id: document.id, title: document.title });
        }
      } catch {
        // An invalid document does not hide the other documents in the project.
      }
    }
  }
  documents.sort((a, b) => a.title.localeCompare(b.title));
  fs.writeFileSync("result.json", JSON.stringify({ documents }), { flag: "wx" });
  return { resultFile: true };
}

function run(job) {
  const root = fs.realpathSync(job.root);
  if (job.type === "list") return listDocuments(root);
  if (job.type === "read") {
    if (!/\.pptx$/i.test(job.source)) throw new Error("Expected a .pptx project file");
    const file = inside(root, job.source);
    const fd = fs.openSync(file, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW || 0));
    try {
      const stat = fs.fstatSync(fd);
      if (!stat.isFile() || stat.size > fileLimit) throw new Error("PPTX exceeds 12 MB");
      const stamp = [stat.ino, stat.size, stat.mtimeMs, stat.ctimeMs].join("/");
      if (job.stamp && job.stamp !== stamp) throw new Error("Source file changed during import");
      if (!Number.isSafeInteger(job.offset) || job.offset < 0 || job.offset >= stat.size)
        throw new Error("Invalid offset or empty source");
      const data = Buffer.allocUnsafe(Math.min(chunkSize, stat.size - job.offset));
      const count = fs.readSync(fd, data, 0, data.length, job.offset);
      const after = fs.fstatSync(fd);
      if (count !== data.length || after.mtimeMs !== stat.mtimeMs || after.ctimeMs !== stat.ctimeMs) {
        throw new Error("Source file changed during import");
      }
      return { data: data.toString("base64"), size: stat.size, stamp };
    } finally {
      fs.closeSync(fd);
    }
  }

  if (!isId(job.documentId)) throw new Error("Invalid document ID");
  if (job.type === "save" || job.type === "load") return documentJob(root, job);
  if (job.type === "asset-read") {
    if (!isId(job.resource)) throw new Error("Invalid resource");
    const asset = inside(root, `presentations/${job.documentId}/resources/${job.resource}.png`);
    if (fs.statSync(asset).size > 2_000_000) throw new Error("Image exceeds 2 MB");
    fs.writeFileSync("result.json", JSON.stringify({ data: fs.readFileSync(asset).toString("base64") }), {
      flag: "wx",
    });
    return { resultFile: true };
  }
  if (job.type === "asset-import") {
    // The agent points at a PNG inside the project; it becomes a resource of the document.
    if (!isId(job.resource)) throw new Error("Invalid resource");
    if (typeof job.source !== "string" || !/\.png$/i.test(job.source))
      throw new Error("Expected a .png project file; convert other formats to PNG first");
    const fd = fs.openSync(inside(root, job.source), fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW || 0));
    let bytes;
    try {
      const stat = fs.fstatSync(fd);
      if (!stat.isFile() || !stat.size || stat.size > 2_000_000) throw new Error("Image must be a PNG of at most 2 MB");
      bytes = Buffer.alloc(stat.size);
      if (fs.readSync(fd, bytes, 0, stat.size, 0) !== stat.size) throw new Error("Source file changed during import");
    } finally {
      fs.closeSync(fd);
    }
    if (bytes.length < 24 || !bytes.subarray(0, 8).equals(pngSignature) || bytes.toString("latin1", 12, 16) !== "IHDR")
      throw new Error("The file is not a PNG image");
    const width = bytes.readUInt32BE(16);
    const height = bytes.readUInt32BE(20);
    if (!width || !height) throw new Error("The file is not a PNG image");
    const target = inside(root, `presentations/${job.documentId}/resources/${job.resource}.png`, true);
    if (fs.existsSync(target)) throw new Error("Image ID already exists");
    fs.writeFileSync(target, bytes, { flag: "wx" });
    return { width, height };
  }
  if (job.type === "asset-remove") {
    if (!isId(job.resource)) throw new Error("Invalid resource");
    fs.rmSync(inside(root, `presentations/${job.documentId}/resources/${job.resource}.png`), { force: true });
    return { removed: true };
  }
  const asset = job.type === "asset-write" && isId(job.resource);
  const preview = job.type === "preview-write" && isId(job.preview);
  const exported = job.type === "write" && ["presentation.pptx", "presentation.pdf"].includes(job.file);
  if ((!asset && !exported && !preview) || !Number.isSafeInteger(job.parts) || job.parts < 1 || job.parts > 342) {
    throw new Error("Invalid export job");
  }
  const target = inside(
    root,
    asset
      ? `presentations/${job.documentId}/resources/${job.resource}.png`
      : preview
        ? `presentations/${job.documentId}/previews/${job.preview}.png`
        : `presentations/${job.documentId}/exports/${job.file}`,
    true,
  );
  if (asset && fs.existsSync(target)) throw new Error("Image ID already exists");
  const temp = path.join(path.dirname(target), `.${randomUUID()}.tmp`);
  const fd = fs.openSync(temp, "wx");
  try {
    try {
      let size = 0;
      for (let i = 0; i < job.parts; i++) {
        const part = fs.readFileSync(inside(process.cwd(), `part-${i}.txt`), "utf8");
        if (part.length > 48 * 1024 || !/^[a-z0-9+/]+={0,2}$/i.test(part)) throw new Error("Invalid export chunk");
        const bytes = Buffer.from(part, "base64");
        if ((asset || preview) && i === 0 && !bytes.subarray(0, 8).equals(pngSignature))
          throw new Error("Expected a PNG resource");
        size += bytes.length;
        if (size > (asset || preview ? 2_000_000 : fileLimit)) throw new Error("Export exceeds 12 MB");
        fs.writeFileSync(fd, bytes);
      }
    } finally {
      fs.closeSync(fd);
    }
    if (asset) fs.linkSync(temp, target);
    else fs.renameSync(temp, target);
    return { saved: true, ...(preview && { path: `presentations/${job.documentId}/previews/${job.preview}.png` }) };
  } finally {
    fs.rmSync(temp, { force: true });
  }
}

try {
  console.log(JSON.stringify(run(JSON.parse(fs.readFileSync("job.json", "utf8")))));
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
