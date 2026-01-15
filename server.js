import express from "express";
import axios from "axios";
import * as cheerio from "cheerio";
import cors from "cors";

import path from "path";
import fs from "fs";
import XLSX from "xlsx";
import { fileURLToPath } from "url";
import multer from "multer";

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));

// =====================
// Paths / Static
// =====================
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
app.use(express.static(path.join(__dirname, "public")));

// =====================
// ENV
// =====================
const WORKER_URL = (process.env.WORKER_URL || "https://1.doanngocminhquy.workers.dev").replace(/\/+$/, "");
const FREESHIP_KEY = process.env.FREESHIP_KEY || "quy892006";
const EXCEL_PATH = path.join(__dirname, "data", "orders.xlsx");

// =====================
// Excel cache
// =====================
let EXCEL_CACHE = null;
let EXCEL_MTIME = 0;

// =====================
// Autopee headers (quan trọng khi deploy Render)
/// =====================
const AUTOPEE_HEADERS = {
  "Accept": "application/json, text/plain, */*",
  "Content-Type": "application/json",
  "Origin": "https://www.autopee.com",
  "Referer": "https://www.autopee.com/products/shopee/vouchers",
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
};

// =====================
// Helpers
// =====================
function sanitize(s) {
  return String(s || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "")
    .trim();
}

function containsOutOfStockMessage(msg) {
  const t = String(msg || "").toLowerCase();
  return (
    t.includes("hết lượt") ||
    t.includes("hết số lượng") ||
    t.includes("đã hết lượt") ||
    t.includes("fully claimed") ||
    t.includes("out of") && t.includes("stock")
  );
}

function parseSaveResult(code, httpStatus, body) {
  const success = body?.success === true;
  const err = body?.data?.error;
  const msg = body?.data?.error_msg || body?.message || "";

  // ưu tiên nhận diện "hết lượt"
  if (containsOutOfStockMessage(msg)) {
    return {
      code,
      httpStatus,
      success,
      error: err ?? null,
      message: msg || null,
      state: "fully_claimed",
      pretty: "⚠️ Mã đã hết lượt sử dụng",
    };
  }

  const state =
    err === 0 ? "saved" :
    err === 5 ? "already_saved" :
    // nếu 2xx nhưng không có err rõ ràng -> coi như hết lượt theo yêu cầu bạn
    (httpStatus >= 200 && httpStatus < 300 && success) ? "unknown_but_ok" :
    "failed";

  const pretty =
    state === "saved" ? "✅ Đã lưu" :
    state === "already_saved" ? "✅ Bạn đã lưu trước đó" :
    state === "unknown_but_ok" ? "⚠️ Mã đã hết lượt sử dụng" :
    `❌ Lỗi: ${msg || "Không rõ"}`;

  return { code, httpStatus, success, error: err ?? null, message: msg || null, state, pretty };
}

async function autopeeList(listUrl) {
  const r = await axios.get(listUrl, {
    headers: AUTOPEE_HEADERS,
    timeout: 20000,
    validateStatus: () => true,
  });

  if (r.status < 200 || r.status >= 300) {
    return { ok: false, status: r.status, data: r.data, message: "HTTP error" };
  }

  // Autopee format: { success: true, data: [...] }
  if (r.data?.success !== true || !Array.isArray(r.data?.data)) {
    return { ok: false, status: 502, data: r.data, message: "Autopee format invalid" };
  }

  const out = r.data.data.map((x) => ({
    voucherCode: x.voucherCode,
    voucherName: x.voucherName || x.voucherCode,
    promotionId: x.promotionId,
    signature: x.signature,
    description: x.description || "",
    startTime: x.startTime || 0,
    endTime: x.endTime || 0,
    iconText: x.iconText || "",
    voucherMarketType: x.voucherMarketType ?? null,
    hasExpired: !!x.hasExpired,
    fullyClaimed: !!x.fullyClaimed,
    disabled: !!x.disabled,
    discountValue: x.discountValue ?? 0,
    discountCap: x.discountCap ?? 0,
    discountPercentage: x.discountPercentage ?? 0,
    minSpend: x.minSpend ?? 0,
  }));

  return { ok: true, out };
}

async function autopeeSave({ cookie, code, listUrl }) {
  // 1) lấy list (để lấy promotionId + signature)
  const listResp = await axios.get(listUrl, {
    headers: AUTOPEE_HEADERS,
    timeout: 20000,
    validateStatus: () => true,
  });

  if (listResp.status < 200 || listResp.status >= 300) {
    return { ok: false, step: "list", status: listResp.status, message: "List HTTP error" };
  }
  if (listResp.data?.success !== true || !Array.isArray(listResp.data?.data)) {
    return { ok: false, step: "list", status: 502, message: "List format invalid", data: listResp.data };
  }

  const list = listResp.data.data;
  const v = list.find((x) => String(x.voucherCode || "").trim() === String(code || "").trim());
  if (!v) return { ok: false, status: 404, message: "Không tìm thấy voucher trong list", code };

  // 2) gọi save
  const payload = {
    cookie,
    voucher_code: v.voucherCode,
    voucher_promotionid: v.promotionId,
    signature: v.signature,
  };

  const r = await axios.post("https://api.autopee.com/shopee/save-voucher", payload, {
    headers: AUTOPEE_HEADERS,
    timeout: 20000,
    validateStatus: () => true,
  });

  return { ok: true, result: parseSaveResult(code, r.status, r.data) };
}

// ======================================================
// 1) COOKIE MODE: /api/orders  (via Worker -> nganmiu)
// ======================================================
app.post("/api/orders", async (req, res) => {
  try {
    let { cookies } = req.body || {};
    if (!Array.isArray(cookies)) cookies = [cookies];
    cookies = cookies.map(sanitize).filter(Boolean);
    if (!cookies.length) return res.status(400).json({ error: "Chưa có cookie" });

    const r = await axios.post(
      `${WORKER_URL}/orders`,
      { cookies },
      {
        headers: { "Content-Type": "application/json" },
        timeout: 25000,
        responseType: "arraybuffer",
        validateStatus: () => true,
      }
    );

    if (r.status < 200 || r.status >= 300) {
      return res.status(502).json({
        error: "Worker trả lỗi",
        worker_status: r.status,
        detail: Buffer.from(r.data || "").toString("utf8").slice(0, 1500),
      });
    }

    const html = Buffer.from(r.data || "").toString("utf8");
    if (html.trim().startsWith("{")) {
      return res.status(502).json({ error: "Worker trả JSON", detail: html });
    }

    const $ = cheerio.load(html);
    const orders = [];

    $("table tbody tr").each((_, tr) => {
      const tds = $(tr).find("td");
      if (tds.length < 9) return;

      orders.push({
        stt: $(tds[0]).text().trim(),
        productImg:
          $(tds[1]).find("img").attr("src") ||
          $(tds[1]).find("img").attr("data-src") ||
          null,
        cod: $(tds[2]).text().trim(),
        mvd: $(tds[3]).text().trim(),
        status: $(tds[4]).text().trim(),
        receiver: $(tds[5]).text().trim(),
        receiverPhone: $(tds[6]).text().trim(),
        address: $(tds[7]).attr("title")?.trim() || $(tds[7]).text().trim(),
        shipperPhone: $(tds[8]).text().trim(),
      });
    });

    return res.json({ count: orders.length, orders });
  } catch (e) {
    return res.status(500).json({
      error: "Lỗi lấy đơn qua Worker",
      detail: e?.response?.data || e.message,
    });
  }
});

// ======================================================
// 2) EXCEL MODE: /api/track?code=...
// ======================================================
function loadExcel() {
  const stat = fs.statSync(EXCEL_PATH);
  if (EXCEL_CACHE && EXCEL_MTIME === stat.mtimeMs) return EXCEL_CACHE;

  const buf = fs.readFileSync(EXCEL_PATH);
  const wb = XLSX.read(buf, { type: "buffer", cellText: false, cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];

  if (!ws || !ws["!ref"]) {
    EXCEL_CACHE = [];
    EXCEL_MTIME = stat.mtimeMs;
    return EXCEL_CACHE;
  }

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rows = [];

  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const A = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const B = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    const C = ws[XLSX.utils.encode_cell({ r, c: 2 })];
    const D = ws[XLSX.utils.encode_cell({ r, c: 3 })];
    const E = ws[XLSX.utils.encode_cell({ r, c: 4 })];
    const F = ws[XLSX.utils.encode_cell({ r, c: 5 })];
    const G = ws[XLSX.utils.encode_cell({ r, c: 6 })];

    const inputOrder = String(A?.v ?? "").trim().toUpperCase();
    const mvd = String(F?.v ?? "").trim().toUpperCase();
    const shopeeOrder = String(G?.v ?? "").trim().toUpperCase();
    if (!inputOrder && !mvd && !shopeeOrder) continue;

    rows.push({
      inputOrder,
      name: String(B?.v ?? "").trim(),
      address: String(C?.v ?? "").trim(),
      product: String(D?.v ?? "").trim(),
      cod: String(E?.v ?? "").trim(),
      mvd,
      shopeeOrder,
    });
  }

  EXCEL_CACHE = rows;
  EXCEL_MTIME = stat.mtimeMs;
  return EXCEL_CACHE;
}

app.get("/api/track", (req, res) => {
  const code = String(req.query.code || "").trim().toUpperCase();
  if (!code) return res.status(400).json({ error: "Thiếu mã (code)" });

  try {
    const list = loadExcel();
    const found = list.find((x) => x.inputOrder === code || x.mvd === code || x.shopeeOrder === code);
    if (!found) return res.status(404).json({ error: "Không tìm thấy trong Excel" });
    return res.json({ order: found });
  } catch (e) {
    return res.status(500).json({ error: "Lỗi đọc Excel", detail: e.message });
  }
});

// ======================================================
// 3) Upload Excel (admin)
// ======================================================
const upload = multer({ storage: multer.memoryStorage() });

function requireAdmin(req, res, next) {
  const token = req.headers["x-admin-token"];
  if (!process.env.ADMIN_TOKEN) return res.status(500).json({ error: "Missing ADMIN_TOKEN on server" });
  if (token !== process.env.ADMIN_TOKEN) return res.status(401).json({ error: "Unauthorized" });
  next();
}

app.post("/api/upload-excel", requireAdmin, upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Chưa chọn file" });
    if (!req.file.originalname.toLowerCase().endsWith(".xlsx")) {
      return res.status(400).json({ error: "Chỉ nhận file .xlsx" });
    }

    const dataDir = path.join(__dirname, "data");
    if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

    fs.writeFileSync(EXCEL_PATH, req.file.buffer);
    EXCEL_CACHE = null;
    EXCEL_MTIME = 0;

    return res.json({ ok: true, message: "Đã cập nhật Excel" });
  } catch (e) {
    return res.status(500).json({ error: "Upload lỗi", detail: e.message });
  }
});

// ======================================================
// 4) AUTOPEE: LIST vouchers (không cần cookie)
// ======================================================
app.post("/api/vouchers", async (req, res) => {
  try {
    const lim = Number(req.body?.limit || 200);
    const url = `https://api.autopee.com/shopee/vouchers?limit=${lim}`;

    const out = await autopeeList(url);
    if (!out.ok) return res.status(502).json({ ok: false, step: "list", ...out });

    return res.json({ ok: true, count: out.out.length, vouchers: out.out });
  } catch (e) {
    return res.status(500).json({ ok: false, error: "List voucher error", detail: e?.response?.data || e.message });
  }
});

// Save discount voucher (cần cookie)
app.post("/api/save-voucher", async (req, res) => {
  try {
    let { cookies, code } = req.body || {};
    code = String(code || "").trim();
    if (!code) return res.status(400).json({ ok: false, message: "Thiếu voucher code" });

    if (!Array.isArray(cookies)) cookies = [cookies];
    cookies = cookies.map(sanitize).filter(Boolean);
    const cookie = cookies.join("\n").trim();
    if (!cookie) return res.status(400).json({ ok: false, message: "Thiếu cookie" });

    const result = await autopeeSave({
      cookie,
      code,
      listUrl: "https://api.autopee.com/shopee/vouchers?limit=200",
    });

    if (!result.ok) return res.status(result.status || 502).json({ ok: false, ...result });
    return res.json({ ok: true, result: result.result });
  } catch (e) {
    return res.status(500).json({ ok: false, error: "Save voucher error", detail: e?.response?.data || e.message });
  }
});

// ======================================================
// 5) AUTOPEE: LIST freeships (không cần cookie)
// ======================================================
app.post("/api/freeships", async (req, res) => {
  try {
    const lim = Number(req.body?.limit || 200);
    const url = `https://api.autopee.com/shopee/freeships?limit=${lim}`;

    const out = await autopeeList(url);
    if (!out.ok) return res.status(502).json({ ok: false, step: "list", ...out });

    return res.json({ ok: true, count: out.out.length, vouchers: out.out });
  } catch (e) {
    return res.status(500).json({ ok: false, error: "List freeship error", detail: e?.response?.data || e.message });
  }
});

// Save freeship (cần cookie + key)
app.post("/api/save-freeship", async (req, res) => {
  try {
    let { cookies, code, key } = req.body || {};
    code = String(code || "").trim();
    key = String(key || "").trim();
    if (!code) return res.status(400).json({ ok: false, message: "Thiếu voucher code" });

    if (key !== FREESHIP_KEY) return res.status(403).json({ ok: false, message: "Sai KEY freeship" });

    if (!Array.isArray(cookies)) cookies = [cookies];
    cookies = cookies.map(sanitize).filter(Boolean);
    const cookie = cookies.join("\n").trim();
    if (!cookie) return res.status(400).json({ ok: false, message: "Thiếu cookie" });

    const result = await autopeeSave({
      cookie,
      code,
      listUrl: "https://api.autopee.com/shopee/freeships?limit=200",
    });

    if (!result.ok) return res.status(result.status || 502).json({ ok: false, ...result });
    return res.json({ ok: true, result: result.result });
  } catch (e) {
    return res.status(500).json({ ok: false, error: "Save freeship error", detail: e?.response?.data || e.message });
  }
});

// ======================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server running on port:", PORT);
  console.log("Excel path:", EXCEL_PATH);
  console.log("WORKER_URL:", WORKER_URL);
  console.log("FREESHIP_KEY:", FREESHIP_KEY ? "(set)" : "(missing)");
});
