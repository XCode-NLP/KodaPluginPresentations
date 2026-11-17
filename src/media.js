export async function readSource(sessionId, relativePath, bridge) {
  let stamp;
  let buffer;
  for (let offset = 0; !buffer || offset < buffer.length;) {
    const part = await bridge.call("readChunk", { sessionId, source: relativePath, offset, stamp });
    stamp = part.stamp;
    buffer ??= new Uint8Array(part.size);
    const bytes = Uint8Array.from(atob(part.data), (char) => char.charCodeAt(0));
    buffer.set(bytes, offset);
    offset += bytes.length;
  }
  return buffer.buffer;
}

export async function snapshot(element, width, height, transparent = false, scale = 2) {
  // Capture the rendered DOM with its computed styles. External resources and
  // embedded media need a separate flattening pass before this can be generalized.
  const clone = element.cloneNode(true);
  const originalNodes = [element, ...element.querySelectorAll("*")];
  const cloneNodes = [clone, ...clone.querySelectorAll("*")];
  for (const [index, node] of originalNodes.entries()) {
    const styles = getComputedStyle(node);
    for (const property of styles) cloneNodes[index].style?.setProperty(property, styles.getPropertyValue(property));
    const copy = cloneNodes[index];
    if (node instanceof HTMLCanvasElement || (node instanceof HTMLImageElement && node.src.startsWith("blob:"))) {
      if (node instanceof HTMLImageElement) await node.decode();
      const raster = document.createElement("canvas");
      raster.width = node.naturalWidth ?? node.width;
      raster.height = node.naturalHeight ?? node.height;
      if (!raster.width || !raster.height || raster.width * raster.height > 16_000_000) continue;
      raster.getContext("2d").drawImage(node, 0, 0);
      const replacement = new Image();
      replacement.src = raster.toDataURL();
      replacement.style.cssText = copy.style.cssText;
      copy.replaceWith(replacement);
    }
    if (node.matches("video, audio, iframe, object")) copy.remove();
  }
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", width);
  svg.setAttribute("height", height);
  const foreign = document.createElementNS(svg.namespaceURI, "foreignObject");
  foreign.setAttribute("width", "100%");
  foreign.setAttribute("height", "100%");
  foreign.append(clone);
  svg.append(foreign);
  const image = new Image();
  image.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(svg));
  await image.decode();
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(width * scale);
  canvas.height = Math.ceil(height * scale);
  const context = canvas.getContext("2d");
  context.fillStyle = "white";
  if (!transparent) context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas;
}

export async function sendBinary(bridge, params, blob) {
  const limit = params.resource || params.preview ? 2_000_000 : 12 * 1024 * 1024;
  if (!blob.size || blob.size > limit) throw new Error("Output exceeds the file size limit");
  const chunkSize = 36 * 1024;
  const count = Math.ceil(blob.size / chunkSize);
  const { sessionId } = params;
  const token = await bridge.call("startUpload", { ...params, count });
  try {
    for (let index = 0; index < count; index++) {
      const data = new Uint8Array(await blob.slice(index * chunkSize, (index + 1) * chunkSize).arrayBuffer());
      const part = btoa(String.fromCharCode(...data));
      await bridge.call("appendUpload", { sessionId, token, index, part });
    }
    await bridge.call("finishUpload", { sessionId, token });
  } catch (error) {
    await bridge.call("abortUpload", { sessionId, token }).catch(() => {});
    throw error;
  }
}
