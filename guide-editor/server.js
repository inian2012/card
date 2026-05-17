const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const express = require("express");
const dotenv = require("dotenv");
const multer = require("multer");

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3333);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "110578";
const SESSION_SECRET = process.env.SESSION_SECRET || crypto.randomBytes(32).toString("hex");
const DATA_DIR = path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "docs.json");
const UPLOAD_DIR = path.join(DATA_DIR, "uploads");
const PUBLIC_DIR = path.join(__dirname, "public");

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

app.set("trust proxy", 1);
app.use(express.json({ limit: "30mb" }));
app.use(express.static(PUBLIC_DIR));

const upload = multer({
  storage: multer.diskStorage({
    destination: UPLOAD_DIR,
    filename: (req, file, callback) => {
      callback(null, `${crypto.randomUUID()}${path.extname(file.originalname).toLowerCase()}`);
    }
  }),
  limits: { fileSize: 80 * 1024 * 1024 },
  fileFilter: (req, file, callback) => {
    const isDocx = path.extname(file.originalname).toLowerCase() === ".docx";
    callback(isDocx ? null : new Error("只支持 .docx 文件"), isDocx);
  }
});

function createToken() {
  return crypto
    .createHmac("sha256", SESSION_SECRET)
    .update(`guide-admin:${ADMIN_PASSWORD}`)
    .digest("hex");
}

function parseCookies(req) {
  return Object.fromEntries(
    String(req.headers.cookie || "")
      .split(";")
      .map((item) => item.trim())
      .filter(Boolean)
      .map((item) => {
        const index = item.indexOf("=");
        return [item.slice(0, index), decodeURIComponent(item.slice(index + 1))];
      })
  );
}

function requireAdmin(req, res, next) {
  const token = parseCookies(req).guide_auth;
  const expectedToken = createToken();
  const tokenBuffer = Buffer.from(String(token || ""));
  const expectedBuffer = Buffer.from(expectedToken);
  if (tokenBuffer.length === expectedBuffer.length && crypto.timingSafeEqual(tokenBuffer, expectedBuffer)) {
    return next();
  }
  return res.status(401).json({ success: false, message: "需要管理权限" });
}

function isAdminRequest(req) {
  return parseCookies(req).guide_auth === createToken();
}

const SHARE_CODE_CHARS = "abcdefghijklmnopqrstuvwxyz0123456789";

function randomShareCode() {
  return Array.from({ length: 3 }, () => SHARE_CODE_CHARS[crypto.randomInt(SHARE_CODE_CHARS.length)]).join("");
}

function normaliseShareCodes(docs) {
  let changed = false;
  const used = new Set();

  for (const doc of docs) {
    if (!/^[a-z0-9]{3}$/.test(String(doc.shareCode || "")) || used.has(doc.shareCode)) {
      delete doc.shareCode;
      changed = true;
      continue;
    }
    used.add(doc.shareCode);
  }

  const missingDocs = docs
    .filter((doc) => !doc.shareCode)
    .sort((a, b) => String(a.createdAt || "").localeCompare(String(b.createdAt || "")));

  for (const doc of missingDocs) {
    let code = randomShareCode();
    while (used.has(code)) code = randomShareCode();
    doc.shareCode = code;
    used.add(doc.shareCode);
    changed = true;
  }

  return changed;
}

function nextShareCode(docs) {
  normaliseShareCodes(docs);
  const used = new Set(docs.map((doc) => doc.shareCode).filter(Boolean));
  let code = randomShareCode();
  while (used.has(code)) code = randomShareCode();
  return code;
}

function loadDocs() {
  if (!fs.existsSync(DATA_FILE)) return [];
  try {
    const parsed = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    if (!Array.isArray(parsed)) return [];
    if (normaliseShareCodes(parsed)) saveDocs(parsed);
    return parsed;
  } catch (error) {
    return [];
  }
}

function saveDocs(docs) {
  const tmpFile = `${DATA_FILE}.tmp`;
  fs.writeFileSync(tmpFile, JSON.stringify(docs, null, 2));
  fs.renameSync(tmpFile, DATA_FILE);
}

function publicDoc(doc) {
  return {
    id: doc.id,
    shareCode: doc.shareCode,
    title: doc.title,
    type: doc.type || "rich",
    visibility: doc.visibility,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt
  };
}

function findDoc(id) {
  return loadDocs().find((doc) => doc.id === id || doc.shareCode === id);
}

function findDocIndex(docs, id) {
  return docs.findIndex((doc) => doc.id === id || doc.shareCode === id);
}

function removeUpload(fileUrl) {
  if (!fileUrl || !fileUrl.startsWith("/uploads/")) return;
  const filename = path.basename(fileUrl);
  fs.promises.unlink(path.join(UPLOAD_DIR, filename)).catch(() => {});
}

app.get("/health", (req, res) => {
  res.json({ success: true, message: "ok" });
});

app.post("/api/login", (req, res) => {
  const password = String(req.body.password || "");
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ success: false, message: "密码错误" });
  }

  res.cookie("guide_auth", createToken(), {
    httpOnly: true,
    sameSite: "lax",
    secure: req.secure || req.headers["x-forwarded-proto"] === "https",
    maxAge: 1000 * 60 * 60 * 24 * 7
  });
  return res.json({ success: true });
});

app.post("/api/logout", (req, res) => {
  res.clearCookie("guide_auth");
  res.json({ success: true });
});

app.get("/api/me", (req, res) => {
  res.json({ success: true, authed: isAdminRequest(req) });
});

app.get("/api/docs", (req, res) => {
  const docs = loadDocs();
  const visibleDocs = isAdminRequest(req) ? docs : docs.filter((doc) => doc.visibility === "public");
  res.json({ success: true, docs: visibleDocs.map(publicDoc) });
});

app.get("/api/docs/:id", (req, res) => {
  const doc = findDoc(req.params.id);
  if (!doc) return res.status(404).json({ success: false, message: "文章不存在" });
  if (doc.visibility !== "public") {
    return res.json({ success: true, locked: true, title: "未公开文章", message: "作者没公开" });
  }
  return res.json({ success: true, doc });
});

app.get("/api/docs/:id/file", (req, res) => {
  const doc = findDoc(req.params.id);
  if (!doc || doc.type !== "word" || !doc.fileUrl) {
    return res.status(404).json({ success: false, message: "Word 文件不存在" });
  }
  if (doc.visibility !== "public" && !isAdminRequest(req)) {
    return res.status(403).json({ success: false, message: "作者没公开" });
  }

  const filePath = path.join(UPLOAD_DIR, path.basename(doc.fileUrl));
  return res.sendFile(filePath);
});

app.get("/api/admin/docs/:id", requireAdmin, (req, res) => {
  const doc = findDoc(req.params.id);
  if (!doc) return res.status(404).json({ success: false, message: "文章不存在" });
  return res.json({ success: true, doc });
});

app.post("/api/admin/docs", requireAdmin, (req, res) => {
  const now = new Date().toISOString();
  const title = String(req.body.title || "未命名文章").trim() || "未命名文章";
  const docs = loadDocs();
  const doc = {
    id: crypto.randomUUID(),
    shareCode: nextShareCode(docs),
    title,
    type: "rich",
    visibility: req.body.visibility === "public" ? "public" : "private",
    content: req.body.content || "",
    createdAt: now,
    updatedAt: now
  };
  docs.unshift(doc);
  saveDocs(docs);
  res.json({ success: true, doc });
});

app.put("/api/admin/docs/:id", requireAdmin, (req, res) => {
  const docs = loadDocs();
  const index = findDocIndex(docs, req.params.id);
  if (index === -1) return res.status(404).json({ success: false, message: "文章不存在" });

  docs[index] = {
    ...docs[index],
    title: String(req.body.title || docs[index].title).trim() || docs[index].title,
    visibility: req.body.visibility === "public" ? "public" : "private",
    content: docs[index].type === "word" ? docs[index].content || "" : String(req.body.content || ""),
    updatedAt: new Date().toISOString()
  };
  saveDocs(docs);
  res.json({ success: true, doc: docs[index] });
});

app.post("/api/admin/import-word", requireAdmin, upload.single("word"), (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: "请选择 .docx 文件" });

  const docs = loadDocs();
  const now = new Date().toISOString();
  const incomingTitle = String(req.body.title || "").trim();
  const filenameTitle = Buffer.from(req.file.originalname, "latin1")
    .toString("utf8")
    .replace(/\.docx$/i, "");
  const title = incomingTitle || filenameTitle || "未命名 Word";
  const fileUrl = `/uploads/${req.file.filename}`;
  const visibility = req.body.visibility === "public" ? "public" : "private";
  const replaceId = String(req.body.docId || "");
  const index = replaceId ? findDocIndex(docs, replaceId) : -1;

  if (index >= 0) {
    removeUpload(docs[index].fileUrl);
    docs[index] = {
      ...docs[index],
      title,
      type: "word",
      visibility,
      content: "",
      fileUrl,
      originalName: req.file.originalname,
      updatedAt: now
    };
    saveDocs(docs);
    return res.json({ success: true, doc: docs[index] });
  }

  const doc = {
    id: crypto.randomUUID(),
    shareCode: nextShareCode(docs),
    title,
    type: "word",
    visibility,
    content: "",
    fileUrl,
    originalName: req.file.originalname,
    createdAt: now,
    updatedAt: now
  };
  docs.unshift(doc);
  saveDocs(docs);
  return res.json({ success: true, doc });
});

app.delete("/api/admin/docs/:id", requireAdmin, (req, res) => {
  const docs = loadDocs();
  const doc = docs.find((item) => item.id === req.params.id || item.shareCode === req.params.id);
  removeUpload(doc?.fileUrl);
  saveDocs(docs.filter((item) => item.id !== req.params.id && item.shareCode !== req.params.id));
  res.json({ success: true });
});

app.get("*", (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Guide editor running on http://127.0.0.1:${PORT}`);
});
