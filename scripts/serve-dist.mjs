// Minimal static HTTPS server for the production build (dist/), using the Office dev certificates.
// Usage: npm run build && npm run serve:prod         (https://localhost:3000)
//        node scripts/serve-dist.mjs --http           (plain HTTP, for UI preview in a browser only)
import { createServer as createHttps } from "node:https";
import { createServer as createHttp } from "node:http";
import { readFile, stat } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import { extname, join, normalize, resolve } from "node:path";
import { homedir } from "node:os";

const root = resolve(process.cwd(), "dist");
const port = Number(process.env.PORT || 3000);
const useHttp = process.argv.includes("--http");
const types = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".json": "application/json",
  ".png": "image/png", ".xml": "application/xml", ".css": "text/css", ".txt": "text/plain",
};

async function handler(req, res) {
  try {
    const url = new URL(req.url, "http://x");
    let p = normalize(join(root, decodeURIComponent(url.pathname)));
    if (!p.startsWith(root)) throw new Error("forbidden");
    if ((await stat(p)).isDirectory()) p = join(p, "taskpane.html");
    const body = await readFile(p);
    res.writeHead(200, { "Content-Type": types[extname(p)] ?? "application/octet-stream", "Access-Control-Allow-Origin": "*" });
    res.end(body);
  } catch {
    res.writeHead(404);
    res.end("Not found");
  }
}

if (useHttp) {
  createHttp(handler).listen(port, () => console.log(`MathPaste (HTTP preview) on http://localhost:${port}/taskpane.html`));
} else {
  const dir = join(homedir(), ".office-addin-dev-certs");
  const key = join(dir, "localhost.key");
  const cert = join(dir, "localhost.crt");
  if (!existsSync(key) || !existsSync(cert)) {
    console.error("Certificats de développement introuvables. Lancez d'abord : npm run certs");
    process.exit(1);
  }
  createHttps({ key: readFileSync(key), cert: readFileSync(cert) }, handler).listen(port, () =>
    console.log(`MathPaste servi sur https://localhost:${port}/taskpane.html (Ctrl+C pour arrêter)`),
  );
}
