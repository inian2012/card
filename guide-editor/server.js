const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const express = require("express");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3333);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "110578";
const SESSION_SECRET = process.env.SESSION_SECRET || crypto.randomBytes(32).toString("hex");
const DATA_DIR = path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "docs.json");
const PUBLIC_DIR = path.join(__dirname, "public");

fs.mkdirSync(DATA_DIR, { recursive: true });

app.set("trust proxy", 1);
app.use(express.json({ limit: "30mb" }));
app.use(express.static(PUBLIC_DIR));

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

function loadDocs() {
  if (!fs.existsSync(DATA_FILE)) return [];
  try {
    const parsed = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    return Array.isArray(parsed) ? parsed : [];
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
    title: doc.title,
    visibility: doc.visibility,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt
  };
}

function findDoc(id) {
  return loadDocs().find((doc) => doc.id === id);
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
    return res.json({ success: true, locked: true, title: doc.title, message: "作者没公开" });
  }
  return res.json({ success: true, doc });
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
    title,
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
  const index = docs.findIndex((doc) => doc.id === req.params.id);
  if (index === -1) return res.status(404).json({ success: false, message: "文章不存在" });

  docs[index] = {
    ...docs[index],
    title: String(req.body.title || docs[index].title).trim() || docs[index].title,
    visibility: req.body.visibility === "public" ? "public" : "private",
    content: String(req.body.content || ""),
    updatedAt: new Date().toISOString()
  };
  saveDocs(docs);
  res.json({ success: true, doc: docs[index] });
});

app.delete("/api/admin/docs/:id", requireAdmin, (req, res) => {
  const docs = loadDocs();
  saveDocs(docs.filter((doc) => doc.id !== req.params.id));
  res.json({ success: true });
});

app.get("*", (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Guide editor running on http://127.0.0.1:${PORT}`);
});
