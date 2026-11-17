import { build } from "esbuild";
import { copyFile, mkdir, rm } from "node:fs/promises";
import { bundleRunner } from "./scripts/runnerBundle.mjs";
import { writeNotices } from "./scripts/notices.mjs";

await rm("bundle", { recursive: true, force: true });
await mkdir("bundle/ui", { recursive: true });

// The runner is a separate Node script the main part writes into a scratch folder.
const runner = await bundleRunner();
const main = await build({
  entryPoints: ["src/main.ts"],
  metafile: true,
  outfile: "bundle/main.js",
  bundle: true,
  format: "cjs",
  target: "es2022",
  plugins: [
    {
      name: "embedded-runner",
      setup(build) {
        build.onResolve({ filter: /^virtual:runner$/ }, () => ({ path: "runner", namespace: "virtual" }));
        build.onLoad({ filter: /.*/, namespace: "virtual" }, () => ({ contents: runner, loader: "text" }));
      },
    },
  ],
});
const ui = await build({
  entryPoints: ["src/app.ts"],
  metafile: true,
  outfile: "bundle/ui/app.js",
  format: "esm",
  bundle: true,
  minify: true,
  target: "es2022",
  legalComments: "inline",
});
for (const file of ["index.html", "style.css"]) await copyFile(`src/ui/${file}`, `bundle/ui/${file}`);
// The isolated panel cannot inherit styles from the Desktop renderer.
await copyFile("src/tokens.css", "bundle/ui/tokens.css");
writeNotices([...Object.keys(main.metafile.inputs), ...Object.keys(ui.metafile.inputs)], "bundle/ui/THIRD_PARTY_NOTICES.txt");
