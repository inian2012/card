const path = require("path");
const express = require("express");
const axios = require("axios");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;
const BASE = process.env.CARD_API_BASE || "https://card.efuncard.com/api/external";
const TOKEN = process.env.CARD_API_TOKEN;

if (!TOKEN) {
  console.warn("Missing CARD_API_TOKEN. Create a .env file before starting the server.");
}

app.use(express.json());
app.use(express.static(__dirname));

function buildHeaders() {
  return {
    Authorization: `Bearer ${TOKEN}`,
    "Content-Type": "application/json"
  };
}

function getUpstreamErrorMessage(error, fallback) {
  const upstream = error.response && error.response.data;
  if (upstream && typeof upstream === "object") {
    return upstream.message || upstream.error || JSON.stringify(upstream);
  }
  return fallback;
}

app.get("/health", (req, res) => {
  res.json({ success: true, message: "ok" });
});

app.post("/api/query-card", async (req, res) => {
  const code = String(req.body.code || "").trim();

  if (!code) {
    return res.status(400).json({ success: false, message: "缺少 CDK 激活码" });
  }

  if (!TOKEN) {
    return res.status(500).json({ success: false, message: "服务端未配置 CARD_API_TOKEN" });
  }

  let redeemMessage = "未执行开卡";

  try {
    const redeemResponse = await axios.post(
      `${BASE}/redeem`,
      { code },
      { headers: buildHeaders(), timeout: 15000 }
    );
    const redeemData = redeemResponse.data || {};
    redeemMessage = redeemData.success ? "开卡成功" : (redeemData.message || "激活码已使用，直接查询卡片");
  } catch (error) {
    redeemMessage = "开卡接口异常，已直接查询卡片";
  }

  try {
    const cardResponse = await axios.get(
      `${BASE}/cards/query/${encodeURIComponent(code)}`,
      { headers: buildHeaders(), timeout: 15000 }
    );

    const cardData = cardResponse.data || {};
    if (!cardData.success) {
      return res.status(400).json({
        success: false,
        message: cardData.message || "查询失败",
        raw: cardData
      });
    }

    return res.json({
      success: true,
      redeemMessage,
      data: cardData.data
    });
  } catch (error) {
    return res.status(502).json({
      success: false,
      message: getUpstreamErrorMessage(error, "查询卡片失败")
    });
  }
});

app.post("/api/3ds-link", async (req, res) => {
  const code = String(req.body.code || "").trim();
  const minutes = Number(req.body.minutes || 5);

  if (!code) {
    return res.status(400).json({ success: false, message: "缺少 CDK 激活码" });
  }

  if (!TOKEN) {
    return res.status(500).json({ success: false, message: "服务端未配置 CARD_API_TOKEN" });
  }

  try {
    const verifyResponse = await axios.post(
      `${BASE}/3ds/verify`,
      { code, minutes },
      { headers: buildHeaders(), timeout: 15000 }
    );
    const verifyData = verifyResponse.data || {};

    if (!verifyData.success) {
      return res.status(400).json({
        success: false,
        message: verifyData.message || "获取 3DS 链接失败",
        raw: verifyData
      });
    }

    return res.json({
      success: true,
      data: verifyData.data || {}
    });
  } catch (error) {
    return res.status(502).json({
      success: false,
      message: getUpstreamErrorMessage(error, "获取 3DS 链接失败")
    });
  }
});

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
