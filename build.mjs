import { build } from "esbuild";
import { copyFile, readFile, writeFile } from "node:fs/promises";

await build({
  entryPoints: ["src/main.js"],
  outfile: "main.js",
  bundle: true,
  format: "cjs",
  plugins: [
    {
      name: "embedded-runner",
      setup(build) {
        build.onLoad({ filter: /runner\.cjs$/ }, async ({ path }) => ({
          contents: await readFile(path, "utf8"),
          loader: "text",
        }));
      },
    },
  ],
});
await build({
  entryPoints: ["src/app.js"],
  outfile: "ui/app.js",
  format: "esm",
  bundle: true,
  minify: true,
  legalComments: "inline",
});
// The isolated panel cannot inherit styles from the Desktop renderer.
await copyFile(new URL("./src/tokens.css", import.meta.url), "ui/tokens.css");
const notices = await Promise.all(
  [
    "@aiden0z/pptx-renderer/LICENSE",
    "@aiden0z/pptx-renderer/THIRD_PARTY_NOTICES.md",
    "pptxgenjs/LICENSE",
    "jspdf/LICENSE",
  ].map((file) => readFile(`node_modules/${file}`, "utf8")),
);
await writeFile("ui/THIRD_PARTY_NOTICES.txt", notices.join("\n\n"));
