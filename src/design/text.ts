// Text measurement without a canvas: the main part runs in QuickJS, so the width of a line is
// estimated from average glyph widths. The estimate leans wide, so "fits" errs on the safe side.

export const LINE_HEIGHT = 1.16;

const FACTORS: Record<string, number> = {
  arial: 0.52,
  helvetica: 0.52,
  "trebuchet ms": 0.52,
  tahoma: 0.53,
  verdana: 0.6,
  georgia: 0.55,
  "times new roman": 0.47,
  "courier new": 0.61,
  impact: 0.46,
};

function charWidth(font: string, text: string): number {
  const base = FACTORS[String(font).toLowerCase()] ?? 0.55;
  const letters = text.replace(/\s/g, "");
  if (!letters) return base;
  const wide = (letters.match(/[А-ЯЁA-ZШЩЖМЮWMmшщжмю0-9]/g) ?? []).length / letters.length;
  const cyrillic = /[а-яё]/i.test(letters) ? 0.04 : 0;
  return base + cyrillic + wide * 0.08;
}

/** Number of wrapped lines a text takes in a box of the given width. */
export function lineCount(text: string, width: number, size: number, font: string): number {
  const glyph = charWidth(font, text) * size;
  const perLine = Math.max(1, Math.floor(width / glyph));
  let lines = 0;
  for (const paragraph of String(text).split("\n")) {
    let used = 0;
    let count = 1;
    for (const word of paragraph.split(/\s+/).filter(Boolean)) {
      const length = word.length;
      if (length > perLine) {
        // overflow-wrap: anywhere breaks long words.
        count += Math.ceil((used + length) / perLine) - 1;
        used = (used + length) % perLine;
        continue;
      }
      if (used && used + 1 + length > perLine) {
        count++;
        used = length;
      } else used += (used ? 1 : 0) + length;
    }
    lines += count;
  }
  return lines;
}

export const textHeight = (text: string, width: number, size: number, font: string, lineHeight = LINE_HEIGHT): number =>
  lineCount(text, width, size, font) * size * lineHeight;

/** Shrink the size until the text fits the box, not below minSize. */
export function fitText(text: string, w: number, h: number, size: number, font: string, minSize: number): { size: number; fits: boolean } {
  let current = size;
  while (current > minSize && textHeight(text, w, current, font) > h) {
    current = Math.max(minSize, Math.floor(current * 0.93));
  }
  return { size: current, fits: textHeight(text, w, current, font) <= h + 0.5 };
}

export const wordCount = (text: string): number => String(text).split(/\s+/).filter(Boolean).length;
