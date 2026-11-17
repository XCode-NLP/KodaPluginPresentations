// Design tokens. A document stores a resolved copy of its theme, and objects keep the role
// they were given ("accent", "h2", "heading"), so switching the theme repaints the whole deck.

export const COLOR_ROLES = [
  "background",
  "surface",
  "text",
  "muted",
  "line",
  "accent",
  "accent2",
  "onAccent",
  "chart1",
  "chart2",
  "chart3",
  "chart4",
  "chart5",
];
export const FONT_ROLES = ["heading", "body"];
// Sizes relative to a 540-unit-high slide; scaled to the document height.
export const SIZE_SCALE = { display: 72, h1: 46, h2: 34, h3: 24, body: 20, small: 16, caption: 13 };
export const SIZE_ROLES = Object.keys(SIZE_SCALE);
// Below this (in 540-unit terms) text is unreadable from the back of a room.
export const MIN_READABLE = 13;

// Fonts are limited to faces available on macOS, Windows and in PowerPoint without embedding.
const theme = (id, name, mood, use, fonts, c, extra = {}) => ({
  id,
  name,
  mood,
  use,
  fonts,
  colors: {
    background: c[0],
    surface: c[1],
    text: c[2],
    muted: c[3],
    line: c[4],
    accent: c[5],
    accent2: c[6],
    onAccent: c[7],
    chart1: c[5],
    chart2: c[6],
    chart3: c[8],
    chart4: c[9],
    chart5: c[10],
  },
  headingBold: true,
  ...extra,
});

export const THEMES = [
  theme(
    "clean",
    "Clean",
    "Neutral, quiet, lots of white space",
    "Default for product, engineering and internal updates",
    { heading: "Arial", body: "Arial" },
    ["#ffffff", "#f3f4f8", "#1c1f2e", "#5d6276", "#dfe2ea", "#3d5afe", "#00796f", "#ffffff", "#f5a524", "#e5484d", "#8e4ec6"],
  ),
  theme(
    "midnight",
    "Midnight",
    "Dark, technical, high contrast",
    "Developer talks, launches, demos on a big screen",
    { heading: "Verdana", body: "Arial" },
    ["#0e1117", "#181d27", "#f1f3f8", "#9aa3b5", "#2a3140", "#7c8cff", "#3dd6c4", "#0e1117", "#ffc53d", "#ff6b81", "#c18cff"],
  ),
  theme(
    "editorial",
    "Editorial",
    "Magazine-like, serif headlines, warm paper",
    "Strategy, research findings, storytelling",
    { heading: "Georgia", body: "Arial" },
    ["#faf7f2", "#efe9df", "#1f1b16", "#6b6258", "#ddd4c6", "#ad401c", "#2f5d50", "#ffffff", "#c99a2e", "#5b6c8f", "#8a6a9e"],
    { headingBold: false },
  ),
  theme(
    "corporate",
    "Corporate",
    "Trustworthy navy, restrained",
    "Board, finance, partners, sales decks",
    { heading: "Arial", body: "Arial" },
    ["#ffffff", "#eef2f7", "#0f2340", "#56657a", "#d6dde8", "#1456b8", "#08795d", "#ffffff", "#e0a100", "#c2414b", "#6e5bb8"],
  ),
  theme(
    "vivid",
    "Vivid",
    "Bold, energetic, saturated color",
    "Marketing, pitches, events, education for young audiences",
    { heading: "Trebuchet MS", body: "Trebuchet MS" },
    ["#ffffff", "#fff3e8", "#22123b", "#6a5a80", "#f0e3f5", "#d6204c", "#6c3cff", "#ffffff", "#ffb020", "#00b7c7", "#2fbf71"],
  ),
  theme(
    "forest",
    "Forest",
    "Calm, natural, green",
    "Sustainability, health, HR, culture",
    { heading: "Georgia", body: "Verdana" },
    ["#f7f9f4", "#e8efe3", "#1d2a1f", "#5c6b5e", "#d3ddcd", "#2a744a", "#915c1e", "#ffffff", "#7a9e3b", "#3f6f8f", "#a85a4a"],
  ),
  theme(
    "mono",
    "Mono",
    "Black and white with one sharp accent",
    "Design reviews, minimal keynotes, manifestos",
    { heading: "Arial", body: "Arial" },
    ["#ffffff", "#f2f2f2", "#111111", "#666666", "#dddddd", "#cc2a10", "#111111", "#ffffff", "#999999", "#444444", "#bbbbbb"],
  ),
  theme(
    "ocean",
    "Ocean",
    "Fresh, light blue, friendly",
    "Onboarding, customer success, healthcare, travel",
    { heading: "Tahoma", body: "Verdana" },
    ["#f6fbfd", "#e3f2f8", "#0d2b3a", "#4f6b78", "#cfe5ee", "#0872a0", "#a94d19", "#ffffff", "#20b486", "#6b5bd6", "#f2c12e"],
  ),
  theme(
    "terminal",
    "Terminal",
    "Dark with monospace accents, hacker aesthetic",
    "CLI tools, infrastructure, security, code walkthroughs",
    { heading: "Courier New", body: "Verdana" },
    ["#101410", "#1a211a", "#e6f0e6", "#8fa38f", "#2c382c", "#4cd964", "#ffcc33", "#101410", "#5ac8fa", "#ff6b6b", "#bf8cff"],
  ),
  theme(
    "sunset",
    "Sunset",
    "Warm dark plum with coral accent",
    "Vision talks, keynotes, product stories",
    { heading: "Georgia", body: "Trebuchet MS" },
    ["#1e1428", "#2b1f38", "#fbeee6", "#c7b3c4", "#3d2e4c", "#ff7a59", "#ffc15e", "#1e1428", "#8fd3c1", "#c38cff", "#ff5d8f"],
    { headingBold: false },
  ),
];

export const DEFAULT_THEME = "clean";

export const themeById = (id) => THEMES.find((item) => item.id === id);

/** The theme stored in a document: tokens only, without catalog prose. */
export function resolvedTheme(id, overrides = {}) {
  const base = themeById(id);
  if (!base) throw new Error(`Unknown theme: ${id}. Call get_design_system for the list.`);
  const colors = { ...base.colors, ...overrides.colors };
  const fonts = { ...base.fonts, ...overrides.fonts };
  return { id: base.id, background: colors.background, colors, fonts, headingBold: base.headingBold };
}

/** Old documents only have a background; give them the clean tokens on that background. */
export function themeOf(doc) {
  const stored = doc.theme ?? {};
  if (stored.colors && stored.fonts) return stored;
  const base = resolvedTheme(DEFAULT_THEME);
  return { ...base, id: null, background: stored.background ?? base.background, colors: { ...base.colors, background: stored.background ?? base.background } };
}

export const sizeFor = (role, height) => Math.round((SIZE_SCALE[role] * height) / 540);

/** Turn a role name or literal into a value. Returns { value, role } where role is null for literals. */
export function resolveToken(field, value, theme, height) {
  if (typeof value !== "string" && field !== "size") return { value, role: null };
  if (field === "size") {
    if (typeof value === "string" && SIZE_SCALE[value]) return { value: sizeFor(value, height), role: value };
    return { value, role: null };
  }
  if (field === "font") {
    if (FONT_ROLES.includes(value)) return { value: theme.fonts[value], role: value };
    return { value, role: null };
  }
  if (COLOR_ROLES.includes(value)) return { value: theme.colors[value], role: value };
  return { value, role: null };
}

export const catalog = () =>
  THEMES.map(({ id, name, mood, use, fonts, colors }) => ({
    id,
    name,
    mood,
    use,
    fonts,
    dark: luminance(colors.background) < 0.3,
    palette: { background: colors.background, text: colors.text, accent: colors.accent, accent2: colors.accent2 },
  }));

export function luminance(hex) {
  const channel = (i) => {
    const v = parseInt(hex.slice(1 + i * 2, 3 + i * 2), 16) / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(0) + 0.7152 * channel(1) + 0.0722 * channel(2);
}

export function contrast(a, b) {
  const [x, y] = [luminance(a), luminance(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}
