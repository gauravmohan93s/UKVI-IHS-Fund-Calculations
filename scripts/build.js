import { build } from "esbuild";
import path from "path";

const root = process.cwd();
const publicDir = path.join(root, "public");

await build({
  entryPoints: [path.join(publicDir, "app.js")],
  outfile: path.join(publicDir, "app.min.js"),
  bundle: true,
  minify: true,
  sourcemap: false,
  target: ["es2018"]
});

await build({
  entryPoints: [path.join(publicDir, "styles.css")],
  outfile: path.join(publicDir, "styles.min.css"),
  bundle: true,
  minify: true,
  sourcemap: false,
  target: ["es2018"],
  external: ["/assets/*"]
});

console.log("Build complete: app.min.js, styles.min.css");
