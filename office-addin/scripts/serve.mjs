import { createServer } from "node:https";
import { readFile, stat } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const devCerts = require("office-addin-dev-certs");

const currentDir = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(currentDir, "../src");
const port = Number(process.env.PORT || 3010);

const mimeTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".svg", "image/svg+xml"],
  [".png", "image/png"],
  [".ico", "image/x-icon"],
]);

function resolveRequestPath(rawUrl) {
  const pathname = decodeURIComponent(new URL(rawUrl, `https://localhost:${port}`).pathname);
  const requested = pathname === "/" ? "/taskpane.html" : pathname;
  const resolved = path.resolve(rootDir, `.${requested}`);
  if (resolved !== rootDir && !resolved.startsWith(`${rootDir}${path.sep}`)) {
    return null;
  }
  return resolved;
}

const httpsOptions = await devCerts.getHttpsServerOptions();

const server = createServer(httpsOptions, async (request, response) => {
  if (request.url === "/health") {
    response.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    response.end(JSON.stringify({ status: "ok", service: "word-addin-ui" }));
    return;
  }

  const filePath = resolveRequestPath(request.url || "/");
  if (!filePath) {
    response.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Invalid path");
    return;
  }

  try {
    const details = await stat(filePath);
    if (!details.isFile()) {
      throw new Error("Not a file");
    }
    const content = await readFile(filePath);
    response.writeHead(200, {
      "Content-Type": mimeTypes.get(path.extname(filePath).toLowerCase()) || "application/octet-stream",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    });
    response.end(content);
  } catch {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Word add-in task pane: https://localhost:${port}/taskpane.html`);
});

function shutdown() {
  server.close(() => process.exit(0));
}

process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
