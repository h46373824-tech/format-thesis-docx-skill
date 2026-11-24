const DOCX_EXTENSION = ".docx";
const MAX_REDIRECTS = 3;

export function assertSafeDownloadUrl(url, base) {
  const baseUrl = new URL(base);
  const target = new URL(url, baseUrl);
  if (target.protocol !== "https:") {
    throw new Error("下载地址必须使用 HTTPS。");
  }
  if (target.username || target.password) {
    throw new Error("下载地址不能包含用户信息。");
  }
  if (target.host !== baseUrl.host) {
    throw new Error("下载地址必须与格式化服务同源。");
  }
  return target;
}

export async function fetchSameOrigin(url, base, fetchImpl, options = {}) {
  let current = assertSafeDownloadUrl(url, base);
  for (let redirectCount = 0; redirectCount <= MAX_REDIRECTS; redirectCount += 1) {
    const response = await fetchImpl(current, {
      ...options,
      redirect: "manual",
    });
    const status = response.status || 0;
    if (status < 300 || status >= 400) {
      return response;
    }
    if (redirectCount === MAX_REDIRECTS) {
      throw new Error("下载重定向次数过多。");
    }
    const location = response.headers?.get?.("location");
    if (!location) {
      throw new Error("下载重定向缺少 Location。");
    }
    current = assertSafeDownloadUrl(location, current);
  }
  throw new Error("下载重定向次数过多。");
}

export function filenameFromDisposition(disposition) {
  const utf8 = String(disposition || "").match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8) {
    return decodeURIComponent(utf8[1].replace(/["']/g, ""));
  }
  const plain = String(disposition || "").match(/filename="?([^";]+)"?/i);
  return plain?.[1] || null;
}

export function safeDocxFilename(name, fallback = `thesis_formatted${DOCX_EXTENSION}`) {
  const base = String(name || "").split(/[/\\]/).filter(Boolean).pop() || "";
  const withoutControl = base.replace(/[\u0000-\u001f]/g, "").trim();
  const stem = withoutControl.replace(/\.[^.]+$/u, "").replace(/[<>:"|?*]/g, "_").trim();
  if (!stem || stem === "." || stem === "..") {
    return fallback;
  }
  return `${stem.slice(0, 120)}${DOCX_EXTENSION}`;
}

export function hasZipMagic(bytes) {
  return bytes.length >= 4
    && bytes[0] === 0x50
    && bytes[1] === 0x4b
    && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07);
}
