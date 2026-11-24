import assert from "node:assert/strict";
import {
  assertSafeDownloadUrl,
  fetchSameOrigin,
  filenameFromDisposition,
  hasZipMagic,
  safeDocxFilename,
} from "../src/download-policy.js";
import { resolveStaticPath } from "./static-path.mjs";

const base = "https://localhost:8787";

assert.equal(
  assertSafeDownloadUrl("https://localhost:8787/jobs/1", base).href,
  "https://localhost:8787/jobs/1",
);
assert.equal(
  assertSafeDownloadUrl("/v1/thesis/result", base).href,
  "https://localhost:8787/v1/thesis/result",
);

for (const blocked of [
  "https://evil.example/file.docx",
  "//evil.example/file.docx",
  "http://localhost:8787/file.docx",
  "https://user:pass@localhost:8787/file.docx",
  "https://169.254.169.254/latest/meta-data/",
  "javascript:alert(1)",
  "file:///etc/passwd",
]) {
  assert.throws(() => assertSafeDownloadUrl(blocked, base), blocked);
}

const calls = [];
const fakeFetch = async (url, options) => {
  calls.push({ href: String(url), redirect: options.redirect });
  if (String(url).endsWith("/start")) {
    return {
      status: 302,
      headers: { get: (name) => (name === "location" ? "/files/result.docx" : null) },
    };
  }
  return { status: 200, headers: { get: () => null } };
};
const ok = await fetchSameOrigin(`${base}/start`, base, fakeFetch, { method: "GET" });
assert.equal(ok.status, 200);
assert.deepEqual(calls.map((call) => call.href), [
  `${base}/start`,
  `${base}/files/result.docx`,
]);
assert.equal(calls[0].redirect, "manual");

await assert.rejects(
  () => fetchSameOrigin(`${base}/start`, base, async () => ({
    status: 302,
    headers: { get: () => "https://evil.example/file.docx" },
  })),
);

await assert.rejects(
  () => fetchSameOrigin(`${base}/loop`, base, async () => ({
    status: 302,
    headers: { get: () => `${base}/loop` },
  })),
);

assert.equal(
  safeDocxFilename(filenameFromDisposition("attachment; filename*=UTF-8''..%2F..%2Fthesis.exe")),
  "thesis.docx",
);
assert.equal(safeDocxFilename("thesis.html"), "thesis.docx");
assert.equal(safeDocxFilename("论文_formatted.docx"), "论文_formatted.docx");
assert.equal(safeDocxFilename(""), "thesis_formatted.docx");

assert.equal(hasZipMagic(Uint8Array.from([0x50, 0x4b, 0x03, 0x04])), true);
assert.equal(hasZipMagic(Uint8Array.from([0x3c, 0x68, 0x74, 0x6d])), false);

const rootDir = "/tmp/thesis-audit-src";
assert.equal(
  resolveStaticPath(rootDir, "/taskpane.html", 3010),
  "/tmp/thesis-audit-src/taskpane.html",
);
assert.equal(resolveStaticPath(rootDir, "/%2e%2e%2f%2e%2e%2fetc/passwd", 3010), null);
assert.equal(resolveStaticPath(rootDir, "/%00taskpane.html", 3010), null);
assert.equal(
  resolveStaticPath(rootDir, "/../secret.txt", 3010),
  "/tmp/thesis-audit-src/secret.txt",
);

console.log("client policy tests passed");
