import path from "node:path";

export function resolveStaticPath(rootDir, rawUrl, port) {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(rawUrl, `https://localhost:${port}`).pathname);
  } catch {
    return null;
  }
  if (pathname.includes("\0")) {
    return null;
  }
  const requested = pathname === "/" ? "/taskpane.html" : pathname;
  const resolved = path.resolve(rootDir, `.${requested}`);
  if (resolved !== rootDir && !resolved.startsWith(`${rootDir}${path.sep}`)) {
    return null;
  }
  return resolved;
}
