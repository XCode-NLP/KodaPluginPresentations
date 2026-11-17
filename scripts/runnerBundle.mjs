// Bundles src/runner.ts into the CommonJS text the main part writes as handler.cjs.
import { build } from "esbuild";

export async function bundleRunner() {
  const result = await build({
    entryPoints: [new URL("../src/runner.ts", import.meta.url).pathname],
    bundle: true,
    platform: "node",
    format: "cjs",
    target: "node22",
    write: false,
  });
  return result.outputFiles[0].text;
}
