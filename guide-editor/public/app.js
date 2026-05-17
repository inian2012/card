const state = {
  authed: false,
  docs: [],
  currentDoc: null,
  quill: null,
  renderId: 0
};

const Font = Quill.import("formats/font");
Font.whitelist = ["microsoft-yahei", "simsun", "simhei", "kaiti", "arial", "times-new-roman", "monospace"];
Quill.register(Font, true);

const SizeStyle = Quill.import("attributors/style/size");
SizeStyle.whitelist = ["12px", "14px", "16px", "18px", "24px", "32px"];
Quill.register(SizeStyle, true);

const docList = document.getElementById("docList");
const readerPane = document.getElementById("readerPane");
const adminBtn = document.getElementById("adminBtn");
const newDocBtn = document.getElementById("newDocBtn");
const loginModal = document.getElementById("loginModal");
const loginForm = document.getElementById("loginForm");
const closeLogin = document.getElementById("closeLogin");
const passwordInput = document.getElementById("passwordInput");
const editorModal = document.getElementById("editorModal");
const editorForm = document.getElementById("editorForm");
const titleInput = document.getElementById("titleInput");
const visibilityInput = document.getElementById("visibilityInput");
const closeEditor = document.getElementById("closeEditor");
const deleteDocBtn = document.getElementById("deleteDocBtn");
const saveState = document.getElementById("saveState");
const importDocBtn = document.getElementById("importDocBtn");
const wordFileInput = document.getElementById("wordFileInput");

const toolbarOptions = [
  [{ font: Font.whitelist }, { size: SizeStyle.whitelist }],
  ["bold", "italic", "underline", "strike"],
  [{ color: [] }, { background: [] }],
  [{ header: 1 }, { header: 2 }],
  [{ list: "ordered" }, { list: "bullet" }],
  [{ align: [] }],
  ["blockquote", "code-block"],
  ["link", "image"],
  ["clean"]
];

function api(url, options = {}) {
  return fetch(url, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options
  }).then(async (res) => {
    const data = await res.json();
    if (!res.ok || data.success === false) throw new Error(data.message || "请求失败");
    return data;
  });
}

function formatTime(value) {
  return value ? new Date(value).toLocaleString("zh-CN", { hour12: false }) : "";
}

function shareUrl(id) {
  return `${location.origin}/doc/${id}`;
}

async function copy(text) {
  await navigator.clipboard.writeText(text);
}

function renderDocs() {
  if (!state.docs.length) {
    docList.innerHTML = `<div class="empty-state">${state.authed ? "还没有文章。" : "还没有公开文章。"}</div>`;
    return;
  }

  docList.innerHTML = state.docs.map((doc) => `
    <div class="doc-card" data-id="${doc.id}">
      <div class="doc-title">${escapeHtml(doc.title)}</div>
      <div class="doc-meta">
        <span>${doc.visibility === "public" ? "公开" : "个人"}</span>
        <span>${formatTime(doc.updatedAt)}</span>
      </div>
    </div>
  `).join("");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function renderArticle(doc, locked = false) {
  if (locked) {
    readerPane.innerHTML = `
      <h1 class="article-title">未公开文章</h1>
      <div class="locked">作者没公开</div>
    `;
    return;
  }

  if (doc.type === "word") {
    readerPane.innerHTML = `
      <h1 class="article-title">${escapeHtml(doc.title)}</h1>
      <div class="doc-meta">
        <span>${doc.visibility === "public" ? "公开" : "个人"}</span>
        <span>更新于 ${formatTime(doc.updatedAt)}</span>
      </div>
      <div class="article-tools">
        ${state.authed ? `<button class="primary" id="editBtn">编辑信息</button>` : ""}
        ${state.authed ? `<button class="ghost" id="copyLinkBtn">复制共享链接</button>` : ""}
      </div>
      <div class="word-status" id="wordStatus">正在加载 Word 文档...</div>
      <div class="word-viewer" id="wordViewer"></div>
    `;

    document.getElementById("editBtn")?.addEventListener("click", () => openEditor(doc));
    document.getElementById("copyLinkBtn")?.addEventListener("click", async () => {
      await copy(shareUrl(doc.id));
      document.getElementById("copyLinkBtn").textContent = "已复制";
    });
    renderWordDocument(doc);
    return;
  }

  readerPane.innerHTML = `
    <h1 class="article-title">${escapeHtml(doc.title)}</h1>
    <div class="doc-meta">
      <span>${doc.visibility === "public" ? "公开" : "个人"}</span>
      <span>更新于 ${formatTime(doc.updatedAt)}</span>
    </div>
    <div class="article-tools">
      ${state.authed ? `<button class="primary" id="editBtn">编辑</button>` : ""}
      ${state.authed ? `<button class="ghost" id="copyLinkBtn">复制共享链接</button>` : ""}
    </div>
    <article class="article-content">${doc.content || "<p>暂无内容</p>"}</article>
  `;

  document.getElementById("editBtn")?.addEventListener("click", () => openEditor(doc));
  document.getElementById("copyLinkBtn")?.addEventListener("click", async () => {
    await copy(shareUrl(doc.id));
    document.getElementById("copyLinkBtn").textContent = "已复制";
  });
}

async function renderWordDocument(doc) {
  const renderId = ++state.renderId;
  const statusEl = document.getElementById("wordStatus");
  const viewerEl = document.getElementById("wordViewer");
  if (!viewerEl || !statusEl) return;

  try {
    if (!window.docx || !window.JSZip) {
      throw new Error("Word 渲染组件加载失败");
    }
    const res = await fetch(`/api/docs/${doc.id}/file`);
    if (!res.ok) throw new Error("Word 文件加载失败");
    const buffer = await res.arrayBuffer();
    if (renderId !== state.renderId) return;
    viewerEl.innerHTML = "";
    await window.docx.renderAsync(buffer, viewerEl, null, {
      className: "docx",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      breakPages: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true
    });
    statusEl.textContent = "Word 已加载，复杂样式以预览模式显示";
  } catch (error) {
    console.error(error);
    statusEl.textContent = "Word 加载失败，请刷新页面或重新导入 .docx 文件";
  }
}

async function loadDocs() {
  const data = await api("/api/docs");
  state.docs = data.docs;
  renderDocs();
}

async function openDoc(id) {
  const url = state.authed ? `/api/admin/docs/${id}` : `/api/docs/${id}`;
  const data = await api(url);
  if (data.locked) return renderArticle(data, true);
  state.currentDoc = data.doc;
  renderArticle(data.doc);
}

function ensureEditor() {
  if (state.quill) return state.quill;
  state.quill = new Quill("#editor", {
    theme: "snow",
    modules: { toolbar: toolbarOptions },
    placeholder: "开始写内容，可以插入图片、设置字体、字号、颜色和加粗..."
  });

  state.quill.getModule("toolbar").addHandler("image", () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.onchange = () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        const range = state.quill.getSelection(true);
        state.quill.insertEmbed(range.index, "image", reader.result);
      };
      reader.readAsDataURL(file);
    };
    input.click();
  });
  return state.quill;
}

function openEditor(doc = null) {
  const quill = ensureEditor();
  state.currentDoc = doc;
  titleInput.value = doc?.title || "";
  visibilityInput.value = doc?.visibility || "private";
  quill.root.innerHTML = doc?.content || "";
  quill.enable(doc?.type !== "word");
  document.querySelector("#editor .ql-editor").dataset.placeholder =
    doc?.type === "word"
      ? "Word 预览模式不能在这里直接改正文；可以修改标题/权限，或点击导入 Word 重新替换文件。"
      : "开始写内容，可以插入图片、设置字体、字号、颜色和加粗...";
  deleteDocBtn.classList.toggle("hidden", !doc);
  saveState.textContent = "";
  editorModal.classList.add("open");
}

async function saveEditor(event) {
  event.preventDefault();
  saveState.textContent = "保存中...";
  const payload = {
    title: titleInput.value.trim() || "未命名文章",
    visibility: visibilityInput.value,
    content: state.quill.root.innerHTML
  };

  const data = state.currentDoc
    ? await api(`/api/admin/docs/${state.currentDoc.id}`, { method: "PUT", body: JSON.stringify(payload) })
    : await api("/api/admin/docs", { method: "POST", body: JSON.stringify(payload) });

  state.currentDoc = data.doc;
  saveState.textContent = "已保存";
  await loadDocs();
  renderArticle(data.doc);
  editorModal.classList.remove("open");
}

async function importWordFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      alert("目前支持导入 .docx 格式，请先把 Word 另存为 .docx 后再上传。");
      return;
    }

    saveState.textContent = "正在上传 Word...";
    const formData = new FormData();
    formData.append("word", file);
    formData.append("title", titleInput.value.trim() || file.name.replace(/\.docx$/i, ""));
    formData.append("visibility", visibilityInput.value);
    if (state.currentDoc?.id) formData.append("docId", state.currentDoc.id);

    const res = await fetch("/api/admin/import-word", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok || data.success === false) throw new Error(data.message || "导入失败");
    state.currentDoc = data.doc;
    if (!titleInput.value.trim()) {
      titleInput.value = file.name.replace(/\.docx$/i, "");
    }
    saveState.textContent = "Word 已导入";
    editorModal.classList.remove("open");
    await loadDocs();
    renderArticle(data.doc);
  } catch (error) {
    console.error(error);
    alert(error.message || "Word 导入失败，请确认文件没有损坏。");
    saveState.textContent = "导入失败";
  } finally {
    wordFileInput.value = "";
  }
}

async function deleteCurrentDoc() {
  if (!state.currentDoc || !confirm("确定删除这篇文章？")) return;
  await api(`/api/admin/docs/${state.currentDoc.id}`, { method: "DELETE" });
  state.currentDoc = null;
  editorModal.classList.remove("open");
  readerPane.innerHTML = '<div class="empty-state">文章已删除。</div>';
  await loadDocs();
}

async function refreshAuth() {
  const data = await api("/api/me");
  state.authed = data.authed;
  newDocBtn.classList.toggle("hidden", !state.authed);
  adminBtn.textContent = state.authed ? "管理中" : "管理端";
}

docList.addEventListener("click", (event) => {
  const card = event.target.closest(".doc-card");
  if (card) openDoc(card.dataset.id);
});

adminBtn.addEventListener("click", () => {
  if (state.authed) return;
  loginModal.classList.add("open");
  passwordInput.focus();
});

closeLogin.addEventListener("click", () => loginModal.classList.remove("open"));
newDocBtn.addEventListener("click", () => openEditor());
closeEditor.addEventListener("click", () => editorModal.classList.remove("open"));
deleteDocBtn.addEventListener("click", deleteCurrentDoc);
importDocBtn.addEventListener("click", () => wordFileInput.click());
wordFileInput.addEventListener("change", importWordFile);
editorForm.addEventListener("submit", saveEditor);

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  await api("/api/login", { method: "POST", body: JSON.stringify({ password: passwordInput.value }) });
  passwordInput.value = "";
  loginModal.classList.remove("open");
  await refreshAuth();
  await loadDocs();
});

(async function init() {
  await refreshAuth();
  await loadDocs();
  const match = location.pathname.match(/^\/doc\/([^/]+)/);
  if (match) openDoc(match[1]);
})();
