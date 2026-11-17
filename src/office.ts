import {
  buildPresentation,
  parseZipLazyMedia,
  RECOMMENDED_ZIP_LIMITS,
  renderSlide,
} from "@aiden0z/pptx-renderer/browser";
import PptxGenJS from "pptxgenjs";
import { jsPDF } from "jspdf";
import { id, object, validate } from "./model";
import { snapshot } from "./media";
import { themeOf } from "./design/themes";
import { pptxChart } from "./chart";
import type { PicNodeData, PresentationData, ShapeNodeData, SlideData, SlideNode } from "@aiden0z/pptx-renderer/browser";
import type { ImageObject, PresentationDocument, Slide, SlideObject, TextObject } from "./types";

/** The renderer's XML wrapper is not exported by name. */
type Xml = NonNullable<PicNodeData["fill"]>;
type Report = (key: string, detail?: unknown) => void;
type Message = (key: string, detail?: string) => string;
type Box = { x: number; y: number; w: number; h: number };

function canvas2d(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas 2D is unavailable");
  return context;
}

const palette = (xml: Xml | undefined): string | undefined => {
  const value = xml?.child("srgbClr").attr("val");
  return value && /^[\da-f]{6}$/i.test(value) ? `#${value}` : undefined;
};
const toBlob = (canvas: HTMLCanvasElement) =>
  new Promise<Blob>((resolve, reject) =>
    canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error("PNG encoding failed")))),
  );

function geometry(node: SlideNode): Box {
  return { x: node.position.x, y: node.position.y, w: Math.max(1, node.size.w), h: Math.max(1, node.size.h) };
}

function fontAvailable(font: string): boolean {
  const context = canvas2d(document.createElement("canvas"));
  const sample = "Широкие буквы · mmmWWWiiil";
  const name = font.replace(/["\\]/g, "");
  return ["serif", "sans-serif", "monospace"].some((fallback) => {
    context.font = `32px ${fallback}`;
    const baseline = context.measureText(sample).width;
    context.font = `32px "${name}", ${fallback}`;
    return Math.abs(baseline - context.measureText(sample).width) > 0.01;
  });
}

function simpleShape(input: SlideNode, report: Report): SlideObject | null {
  if (input.nodeType !== "shape") return null;
  const node: ShapeNodeData = input;
  if (
    node.rotation ||
    node.flipH ||
    node.flipV ||
    node.customGeometry?.exists() ||
    node.shape3d ||
    ![undefined, "rect", "ellipse", "line"].includes(node.presetGeometry)
  )
    return null;
  if (node.fill?.exists() && !["solidFill", "noFill"].includes(node.fill.localName)) return null;
  if (node.line?.child("prstDash").exists() || node.headEnd || node.tailEnd) return null;
  const props = {
    ...geometry(node),
    fill: palette(node.fill) ?? "transparent",
    stroke: palette(node.line?.child("solidFill")) ?? "transparent",
  };
  const paragraphs = node.textBody?.paragraphs ?? [];
  const runs = paragraphs.flatMap((paragraph) => paragraph.runs);
  const content = paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n");
  if (!content)
    return object(
      node.presetGeometry === "ellipse" ? "ellipse" : node.presetGeometry === "line" ? "line" : "rect",
      props,
    );
  if (runs.some((run) => run.math) || node.presetGeometry === "ellipse") return null;
  const style = runs.find((run) => run.properties?.exists())?.properties;
  const font = style?.child("latin").attr("typeface") || "Arial";
  if (font !== "Arial" && !fontAvailable(font)) report("missingFont", font);
  const fontSize = style?.numAttr("sz");
  const alignment = paragraphs[0]?.properties?.attr("algn");
  if (runs.length > 1 || !style || node.fill?.child("schemeClr").exists()) report("textStyle", node.name);
  return object<TextObject>("text", {
    ...props,
    text: content,
    font,
    size: fontSize ? Math.max(4, Math.min(600, fontSize / 75)) : 28,
    color: palette(style?.child("solidFill")) ?? "#202234",
    align: alignment === "ctr" ? "center" : alignment === "r" ? "right" : "left",
    bold: style?.attr("b") === "1",
    italic: style?.attr("i") === "1",
  });
}

async function originalPicture(presentation: PresentationData, slide: SlideData, input: SlideNode): Promise<Blob | null> {
  if (input.nodeType !== "picture") return null;
  const node: PicNodeData = input;
  if (
    node.rotation ||
    node.flipH ||
    node.flipV ||
    node.isAudio ||
    node.isVideo ||
    node.crop ||
    node.customGeometry?.exists() ||
    node.shape3d ||
    (node.presetGeometry && node.presetGeometry !== "rect")
  )
    return null;
  const rel = node.blipEmbed ? slide.rels.get(node.blipEmbed) : undefined;
  if (!rel || rel.targetMode?.toLowerCase() === "external" || !/\.(png|jpe?g|webp|gif)$/i.test(rel.target)) return null;
  const media = await presentation.mediaResolver?.resolve(rel.target);
  if (!media) return null;
  const image = await createImageBitmap(new Blob([media.data as Uint8Array<ArrayBuffer>]), { resizeQuality: "high" });
  try {
    const canvas = document.createElement("canvas");
    const scale = Math.min(1, 1400 / Math.max(image.width, image.height));
    canvas.width = Math.max(1, Math.round(image.width * scale));
    canvas.height = Math.max(1, Math.round(image.height * scale));
    canvas2d(canvas).drawImage(image, 0, 0, canvas.width, canvas.height);
    return toBlob(canvas);
  } finally {
    image.close();
  }
}

async function flatten(presentation: PresentationData, slide: SlideData, node: SlideNode, stage: HTMLElement) {
  const errors: unknown[] = [];
  const handle = renderSlide(
    presentation,
    { ...slide, nodes: [node], showMasterSp: false, background: undefined },
    {
      externalMedia: "block",
      onNodeError: (_id, error) => errors.push(error),
    },
  );
  stage.replaceChildren(handle.element);
  try {
    await handle.ready;
    handle.element.style.background = "transparent";
    if (errors.length || !handle.element.childElementCount) return null;
    const full = await snapshot(handle.element, presentation.width, presentation.height, true);
    const ctx = canvas2d(full);
    const pixels = ctx.getImageData(0, 0, full.width, full.height).data;
    let left = full.width,
      top = full.height,
      right = 0,
      bottom = 0;
    for (let y = 0; y < full.height; y++) {
      for (let x = 0; x < full.width; x++) {
        if (!pixels[(y * full.width + x) * 4 + 3]) continue;
        left = Math.min(left, x);
        top = Math.min(top, y);
        right = Math.max(right, x + 1);
        bottom = Math.max(bottom, y + 1);
      }
    }
    if (right <= left || bottom <= top) return null;
    const canvas = document.createElement("canvas");
    canvas.width = right - left;
    canvas.height = bottom - top;
    canvas2d(canvas).drawImage(full, left, top, canvas.width, canvas.height, 0, 0, canvas.width, canvas.height);
    return {
      blob: await toBlob(canvas),
      props: { x: left / 2, y: top / 2, w: Math.max(1, canvas.width / 2), h: Math.max(1, canvas.height / 2) },
    };
  } finally {
    handle.dispose();
    stage.replaceChildren();
  }
}

export async function importDeck(
  buffer: ArrayBuffer,
  title: string,
  stage: HTMLElement,
  saveAsset: (documentId: string, resource: string, blob: Blob) => Promise<void>,
  message: Message,
): Promise<PresentationDocument> {
  const files = await parseZipLazyMedia(buffer, RECOMMENDED_ZIP_LIMITS);
  const source = buildPresentation(files);
  const doc: PresentationDocument = {
    format: 1,
    id: id(),
    title,
    width: source.width,
    height: source.height,
    theme: { background: "#ffffff" },
    slides: [],
  };
  if (
    !source.slides.length ||
    source.slides.length > 100 ||
    source.width < 16 ||
    source.height < 16 ||
    source.width > 4000 ||
    source.height > 4000 ||
    source.width * source.height > 4_000_000
  )
    throw new Error(message("importLimit"));
  for (const slide of source.slides) {
    if (slide.nodes.length > 250) throw new Error(message("importLimit"));
    const warnings: string[] = [];
    const next: Slide = { id: id(), objects: [], warnings };
    const report: Report = (key, detail) => {
      if (warnings.length < 100) warnings.push(message(key, String(detail || "").slice(0, 130)));
    };
    const layoutPath = source.slideToLayout.get(slide.index) ?? slide.layoutIndex;
    const layout = source.layouts.get(layoutPath);
    const masterPath = source.layoutToMaster.get(layoutPath);
    const master = masterPath ? source.masters.get(masterPath) : undefined;
    const background = slide.background ?? layout?.background ?? master?.background;
    const backgroundColor = palette(background?.child("bgPr").child("solidFill"));
    if (backgroundColor) next.background = backgroundColor;
    const hasTemplateShapes = (template: { spTree: Xml } | undefined) =>
      !!template?.spTree
        .allChildren()
        .some(({ localName }) => ["sp", "pic", "cxnSp", "grpSp", "graphicFrame"].includes(localName));
    if (
      (background?.exists() && !backgroundColor) ||
      (slide.showMasterSp && (hasTemplateShapes(layout) || hasTemplateShapes(master)))
    )
      report("template");
    for (const node of slide.nodes) {
      if (node.hidden) continue;
      const editable = simpleShape(node, report);
      if (editable) {
        next.objects.push(editable);
        continue;
      }
      const name = node.name || node.nodeType;
      const picture = node.nodeType === "picture" ? node : null;
      const pictureKey = picture?.blipEmbed ?? picture?.blipLink;
      const pictureRel = pictureKey ? slide.rels.get(pictureKey) : undefined;
      const externalPicture = pictureRel?.targetMode?.toLowerCase() === "external";
      if (!externalPicture && !picture?.isAudio && !picture?.isVideo && (node.nodeType as string) !== "unknown") {
        const original = await originalPicture(source, slide, node).catch(() => null);
        const image = original
          ? { blob: original, props: geometry(node) }
          : await flatten(source, slide, node, stage).catch(() => null);
        if (image) {
          const resource = id();
          await saveAsset(doc.id, resource, image.blob);
          next.objects.push(object<ImageObject>("image", { ...image.props, fill: "transparent", resource }));
          if (!original) report("flattened", name);
          continue;
        }
      }
      report("unsupported", name);
      const box = geometry(node);
      next.objects.push(
        object<TextObject>("text", {
          ...box,
          w: Math.max(140, box.w),
          h: Math.max(34, box.h),
          text: message("unsupported", name),
          color: "#9d2244",
          fill: "#ffe3ea",
          size: 16,
        }),
      );
    }
    doc.slides.push(next);
  }
  return validate(doc);
}

const hex = (value: string) => value.slice(1);
const solid = (value: string) => value !== "transparent";

export async function exportPowerPoint(doc: PresentationDocument, imageFor: (resource: string) => Promise<string>): Promise<Blob> {
  const pptx = new PptxGenJS();
  const px = (value: number) => value / 96;
  pptx.defineLayout({ name: "KODA", width: px(doc.width), height: px(doc.height) });
  pptx.layout = "KODA";
  pptx.title = doc.title;
  for (const slide of doc.slides) {
    const page = pptx.addSlide();
    const background = slide.background ?? doc.theme.background;
    if (solid(background)) page.background = { color: hex(background) };
    if (slide.notes) page.addNotes(slide.notes);
    for (const obj of slide.objects) {
      const box = { x: px(obj.x), y: px(obj.y), w: px(obj.w), h: px(obj.h) };
      const fill = solid(obj.fill) ? { color: hex(obj.fill) } : { color: "FFFFFF", transparency: 100 };
      const line = solid(obj.stroke) ? { color: hex(obj.stroke), width: 1 } : { color: "FFFFFF", transparency: 100 };
      if (obj.kind === "text") {
        page.addText(obj.text, {
          ...box,
          fill,
          line,
          margin: 0,
          valign: "top",
          lineSpacingMultiple: 1.16,
          fontFace: obj.font,
          fontSize: obj.size * 0.75,
          align: obj.align,
          bold: obj.bold,
          italic: obj.italic,
          color: solid(obj.color) ? hex(obj.color) : "202234",
        });
      } else if (obj.kind === "chart") {
        const chart = pptxChart(pptx, obj, themeOf(doc), box);
        page.addChart(chart.type, chart.data, chart.options);
      } else if (obj.kind === "image") {
        page.addImage({ ...box, data: await imageFor(obj.resource) });
      } else {
        const shape = { rect: pptx.ShapeType.rect, ellipse: pptx.ShapeType.ellipse, line: pptx.ShapeType.line }[obj.kind];
        page.addShape(shape, { ...box, fill, line });
      }
    }
  }
  return (await pptx.write({ outputType: "blob" })) as Blob;
}

export async function exportPdf(
  doc: PresentationDocument,
  slideElement: (slide: Slide) => Promise<HTMLElement>,
  stage: HTMLElement,
): Promise<Blob> {
  const format: [number, number] = [doc.width, doc.height];
  const orientation = doc.width >= doc.height ? "landscape" : "portrait";
  const pdf = new jsPDF({ format, orientation, unit: "px", hotfixes: ["px_scaling"], compress: true });
  for (const [index, slide] of doc.slides.entries()) {
    const element = await slideElement(slide);
    stage.replaceChildren(element);
    try {
      await Promise.all([...element.querySelectorAll("img")].map((img) => img.decode()));
      const canvas = await snapshot(element, ...format);
      if (index) pdf.addPage(format, orientation);
      pdf.addImage(canvas.toDataURL("image/jpeg", 0.95), "JPEG", 0, 0, ...format);
    } finally {
      stage.replaceChildren();
    }
  }
  return pdf.output("blob");
}
