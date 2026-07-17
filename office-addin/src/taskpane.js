/* global Office, Word */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const state = {
  analysisId: null,
  questions: [],
  documentSignals: null,
};

const elements = {};

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.getElementById("unsupported").hidden = false;
    return;
  }

  bindElements();
  bindEvents();
  restoreBackendUrl();
  document.getElementById("app").hidden = false;
  setStatus("ready", "准备就绪", "任务窗格已连接到当前 Word 文档。");
});

function bindElements() {
  const ids = [
    "backend-url",
    "api-token",
    "test-connection",
    "authority-files",
    "file-list",
    "cover-title",
    "cover-author",
    "cover-student-id",
    "cover-school",
    "cover-major",
    "cover-supervisor",
    "cover-degree",
    "cover-date",
    "add-toc",
    "preview-document",
    "analyze-document",
    "format-document",
    "document-summary",
    "confirmation-section",
    "confirmation-list",
    "status-title",
    "status-message",
  ];

  for (const id of ids) {
    elements[id] = document.getElementById(id);
  }
  elements.statusPanel = document.querySelector(".status-panel");
  elements.actionButtons = Array.from(document.querySelectorAll("button"));
}

function bindEvents() {
  elements["backend-url"].addEventListener("change", saveBackendUrl);
  elements["authority-files"].addEventListener("change", renderSelectedFiles);
  elements["test-connection"].addEventListener("click", testConnection);
  elements["preview-document"].addEventListener("click", previewDocument);
  elements["analyze-document"].addEventListener("click", analyzeDocument);
  elements["format-document"].addEventListener("click", formatDocument);
}

function restoreBackendUrl() {
  const saved = localStorage.getItem("format-thesis-backend-url");
  if (saved) {
    elements["backend-url"].value = saved;
  }
}

function saveBackendUrl() {
  const value = elements["backend-url"].value.trim();
  if (value) {
    localStorage.setItem("format-thesis-backend-url", value);
  }
}

function renderSelectedFiles() {
  const files = Array.from(elements["authority-files"].files || []);
  elements["file-list"].replaceChildren();
  for (const file of files) {
    const row = document.createElement("div");
    row.className = "file-chip";
    const name = document.createElement("span");
    name.textContent = file.name;
    const size = document.createElement("span");
    size.textContent = formatBytes(file.size);
    row.append(name, size);
    elements["file-list"].append(row);
  }
}

async function testConnection() {
  await runAction(async () => {
    setStatus("busy", "正在测试连接", "检查格式化服务的健康状态……");
    const response = await backendFetch("/health", { method: "GET" });
    if (!response.ok) {
      throw new Error(`服务返回 HTTP ${response.status}`);
    }
    const details = await readJsonSafely(response);
    setStatus(
      "success",
      "连接成功",
      details?.service ? `已连接：${details.service}` : "格式化服务可以访问。",
    );
  });
}

async function previewDocument() {
  await runAction(async () => {
    setStatus("busy", "正在预检文档", "读取段落样式并导出当前 DOCX 的临时副本……");
    const [signals, documentFile] = await Promise.all([
      collectDocumentSignals(),
      getCurrentDocumentFile(),
    ]);
    state.documentSignals = signals;
    elements["document-summary"].textContent = [
      `文件：${documentFile.name}`,
      `导出大小：${formatBytes(documentFile.size)}`,
      `段落：${signals.paragraph_count}`,
      `识别到的标题段落：${signals.heading_candidates.length}`,
      `当前选择：${signals.selection_text || "（无）"}`,
    ].join("\n");
    elements["document-summary"].hidden = false;
    setStatus("success", "预检完成", "当前文档可导出为 DOCX，并可提交给格式化后端。");
  });
}

async function analyzeDocument() {
  await runAction(async () => {
    ensureAuthorityProvided();
    setStatus("busy", "正在分析格式", "导出论文并发送模板、封面信息和格式选项……");
    const payload = await buildFormData(false);
    const response = await backendFetch("/v1/thesis/analyze", {
      method: "POST",
      body: payload,
    });
    const result = await expectJson(response);
    state.analysisId = result.analysis_id || result.job_id || null;
    renderQuestions(result.questions || []);
    renderBackendSummary(result.summary);

    if (state.questions.length) {
      setStatus("warning", "需要确认", `发现 ${state.questions.length} 个会影响结果的问题。`);
    } else {
      setStatus("success", "分析完成", "没有需要人工确认的关键歧义，可以开始处理。");
    }
  });
}

async function formatDocument() {
  await runAction(async () => {
    ensureAuthorityProvided();
    const confirmations = collectConfirmations();
    setStatus("busy", "正在生成新论文", "格式化服务正在应用样式、分节、页码和目录……");
    const payload = await buildFormData(true, confirmations);
    const response = await backendFetch("/v1/thesis/format", {
      method: "POST",
      body: payload,
    });

    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      const result = await expectJson(response);
      if (Array.isArray(result.questions) && result.questions.length) {
        state.analysisId = result.analysis_id || result.job_id || state.analysisId;
        renderQuestions(result.questions);
        setStatus("warning", "处理已暂停", "请回答新的确认问题后再次点击处理。");
        return;
      }
      if (result.file_base64) {
        downloadBlob(base64ToBlob(result.file_base64, DOCX_MIME), result.filename || outputFilename());
        setStatus("success", "处理完成", "新的 DOCX 已开始下载，原稿未被覆盖。");
        return;
      }
      if (result.download_url) {
        const fileResponse = await backendFetchAbsolute(result.download_url, { method: "GET" });
        if (!fileResponse.ok) {
          throw new Error(`下载结果失败：HTTP ${fileResponse.status}`);
        }
        downloadBlob(await fileResponse.blob(), result.filename || outputFilename());
        setStatus("success", "处理完成", "新的 DOCX 已开始下载，原稿未被覆盖。");
        return;
      }
      throw new Error("服务返回了 JSON，但没有文件或下载地址。");
    }

    if (!response.ok) {
      throw new Error(`格式化服务返回 HTTP ${response.status}`);
    }
    const disposition = response.headers.get("content-disposition") || "";
    downloadBlob(await response.blob(), filenameFromDisposition(disposition) || outputFilename());
    setStatus("success", "处理完成", "新的 DOCX 已开始下载，原稿未被覆盖。");
  });
}

async function buildFormData(includeConfirmations, confirmations = {}) {
  const [documentFile, signals] = await Promise.all([
    getCurrentDocumentFile(),
    state.documentSignals ? Promise.resolve(state.documentSignals) : collectDocumentSignals(),
  ]);
  state.documentSignals = signals;

  const form = new FormData();
  form.append("manuscript", documentFile, documentFile.name);
  for (const file of Array.from(elements["authority-files"].files || [])) {
    form.append("authorities", file, file.name);
  }
  form.append("cover_metadata", JSON.stringify(coverMetadata()));
  form.append(
    "options",
    JSON.stringify({
      add_toc: elements["add-toc"].checked,
      preserve_content: true,
      output_mode: "new_copy",
      document_signals: signals,
    }),
  );
  if (state.analysisId) {
    form.append("analysis_id", state.analysisId);
  }
  if (includeConfirmations) {
    form.append("confirmations", JSON.stringify(confirmations));
  }
  return form;
}

function coverMetadata() {
  return {
    title: elements["cover-title"].value.trim(),
    author: elements["cover-author"].value.trim(),
    student_id: elements["cover-student-id"].value.trim(),
    school: elements["cover-school"].value.trim(),
    major: elements["cover-major"].value.trim(),
    supervisor: elements["cover-supervisor"].value.trim(),
    degree: elements["cover-degree"].value.trim(),
    date: elements["cover-date"].value.trim(),
  };
}

async function collectDocumentSignals() {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const selection = context.document.getSelection();
    paragraphs.load("items/text,style,styleBuiltIn");
    selection.load("text");
    await context.sync();

    const headingCandidates = [];
    for (let index = 0; index < paragraphs.items.length; index += 1) {
      const paragraph = paragraphs.items[index];
      const style = String(paragraph.styleBuiltIn || paragraph.style || "");
      if (/heading|标题/i.test(style)) {
        headingCandidates.push({
          index,
          text: paragraph.text.slice(0, 300),
          style,
        });
      }
      if (headingCandidates.length >= 200) {
        break;
      }
    }

    return {
      paragraph_count: paragraphs.items.length,
      heading_candidates: headingCandidates,
      selection_text: selection.text.slice(0, 500),
    };
  });
}

function getCurrentDocumentFile() {
  return new Promise((resolve, reject) => {
    if (!Office.context.document?.getFileAsync) {
      reject(new Error("当前 Word 客户端不支持导出整个文档。"));
      return;
    }

    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 65536 },
      async (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message || "无法导出当前 Word 文档。"));
          return;
        }

        const file = result.value;
        try {
          const chunks = [];
          let totalLength = 0;
          for (let index = 0; index < file.sliceCount; index += 1) {
            const slice = await getFileSlice(file, index);
            const bytes = slice.data instanceof ArrayBuffer
              ? new Uint8Array(slice.data)
              : Uint8Array.from(slice.data);
            chunks.push(bytes);
            totalLength += bytes.length;
          }

          const output = new Uint8Array(totalLength);
          let offset = 0;
          for (const chunk of chunks) {
            output.set(chunk, offset);
            offset += chunk.length;
          }
          resolve(new File([output], currentDocumentFilename(), { type: DOCX_MIME }));
        } catch (error) {
          reject(error);
        } finally {
          file.closeAsync(() => {});
        }
      },
    );
  });
}

function getFileSlice(file, index) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || `无法读取文档切片 ${index}。`));
      }
    });
  });
}

function currentDocumentFilename() {
  const url = Office.context.document.url || "";
  const candidate = decodeURIComponent(url.split(/[\\/]/).pop() || "").trim();
  if (candidate.toLowerCase().endsWith(".docx")) {
    return candidate;
  }
  return "thesis.docx";
}

function outputFilename() {
  return currentDocumentFilename().replace(/\.docx$/i, "_formatted.docx");
}

function ensureAuthorityProvided() {
  if (!elements["authority-files"].files?.length) {
    throw new Error("请至少选择一个学校模板、合格范文或格式说明文件。");
  }
}

function backendBaseUrl() {
  const raw = elements["backend-url"].value.trim().replace(/\/+$/, "");
  if (!raw) {
    throw new Error("请填写格式化服务地址。");
  }
  const url = new URL(raw);
  if (url.protocol !== "https:") {
    throw new Error("Office 任务窗格要求使用 HTTPS 后端地址。");
  }
  return url.toString().replace(/\/$/, "");
}

function requestHeaders() {
  const headers = {};
  const token = elements["api-token"].value.trim();
  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }
  return headers;
}

function backendFetch(path, options) {
  return fetch(`${backendBaseUrl()}${path}`, {
    ...options,
    headers: { ...requestHeaders(), ...(options.headers || {}) },
  });
}

function backendFetchAbsolute(url, options) {
  const target = new URL(url, backendBaseUrl());
  if (target.protocol !== "https:") {
    throw new Error("下载地址必须使用 HTTPS。");
  }
  return fetch(target, {
    ...options,
    headers: { ...requestHeaders(), ...(options.headers || {}) },
  });
}

async function expectJson(response) {
  const payload = await readJsonSafely(response);
  if (!response.ok) {
    throw new Error(payload?.message || payload?.error || `服务返回 HTTP ${response.status}`);
  }
  if (!payload) {
    throw new Error("服务没有返回有效 JSON。");
  }
  return payload;
}

async function readJsonSafely(response) {
  try {
    return await response.json();
  } catch {
    return null;
  }
}

function renderBackendSummary(summary) {
  if (!summary) {
    return;
  }
  elements["document-summary"].textContent = typeof summary === "string"
    ? summary
    : JSON.stringify(summary, null, 2);
  elements["document-summary"].hidden = false;
}

function renderQuestions(questions) {
  state.questions = Array.isArray(questions) ? questions : [];
  elements["confirmation-list"].replaceChildren();
  elements["confirmation-section"].hidden = state.questions.length === 0;

  state.questions.forEach((question, index) => {
    const id = String(question.id || `question-${index + 1}`);
    const item = document.createElement("article");
    item.className = "confirmation-item";
    item.dataset.questionId = id;

    const title = document.createElement("h3");
    title.textContent = question.prompt || question.question || `确认项 ${index + 1}`;
    item.append(title);

    for (const [label, value] of [
      ["依据", question.evidence],
      ["建议", question.proposed_default],
      ["影响", question.impact],
    ]) {
      if (value) {
        const paragraph = document.createElement("p");
        paragraph.textContent = `${label}：${value}`;
        item.append(paragraph);
      }
    }

    let control;
    if (Array.isArray(question.options) && question.options.length) {
      control = document.createElement("select");
      const placeholder = document.createElement("option");
      placeholder.value = "";
      placeholder.textContent = "请选择";
      control.append(placeholder);
      for (const option of question.options) {
        const optionElement = document.createElement("option");
        optionElement.value = String(option.value ?? option.label ?? option);
        optionElement.textContent = String(option.label ?? option.value ?? option);
        if (option.recommended) {
          optionElement.textContent += "（建议）";
        }
        control.append(optionElement);
      }
    } else {
      control = document.createElement("textarea");
      control.placeholder = "请输入你的决定";
    }
    control.dataset.confirmationInput = id;
    control.dataset.required = question.required === false ? "false" : "true";
    item.append(control);
    elements["confirmation-list"].append(item);
  });
}

function collectConfirmations() {
  const answers = {};
  const missing = [];
  for (const input of elements["confirmation-list"].querySelectorAll("[data-confirmation-input]")) {
    const value = input.value.trim();
    if (!value && input.dataset.required !== "false") {
      missing.push(input.dataset.confirmationInput);
    }
    if (value) {
      answers[input.dataset.confirmationInput] = value;
    }
  }
  if (missing.length) {
    throw new Error(`请先回答 ${missing.length} 个必填确认问题。`);
  }
  return answers;
}

function filenameFromDisposition(disposition) {
  const utf8 = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8) {
    return decodeURIComponent(utf8[1].replace(/["']/g, ""));
  }
  const plain = disposition.match(/filename="?([^";]+)"?/i);
  return plain?.[1] || null;
}

function base64ToBlob(value, mimeType) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mimeType });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.style.display = "none";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function setStatus(stateName, title, message) {
  elements.statusPanel.dataset.state = stateName;
  elements["status-title"].textContent = title;
  elements["status-message"].textContent = message;
}

function setBusy(isBusy) {
  for (const button of elements.actionButtons) {
    button.disabled = isBusy;
  }
}

async function runAction(action) {
  setBusy(true);
  try {
    await action();
  } catch (error) {
    console.error(error);
    setStatus("error", "操作失败", error instanceof Error ? error.message : String(error));
  } finally {
    setBusy(false);
  }
}

function formatBytes(value) {
  if (value < 1024) return `${value} B`;
  if (value < 1024 ** 2) return `${(value / 1024).toFixed(1)} KB`;
  return `${(value / 1024 ** 2).toFixed(1)} MB`;
}
