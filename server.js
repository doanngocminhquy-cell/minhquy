import express from "express";
import axios from "axios";
import * as cheerio from "cheerio";
import cors from "cors";

import path from "path";
import fs from "fs";
import XLSX from "xlsx";
import { fileURLToPath } from "url";

const app = express();
app.use(cors());
app.use(express.json());

// ===== Paths =====
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// serve static
app.use(express.static(path.join(__dirname, "public")));

// ===== Worker URL (KHÔNG CÓ / CUỐI) =====
const WORKER_URL = "https://1.doanngocminhquy.workers.dev";

function sanitize(s) {
  return String(s || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "")
    .trim();
}

// ======================================================
// 1) COOKIE MODE: /api/orders
// ======================================================
app.post("/api/orders", async (req, res) => {
  try {
    let { cookies } = req.body;
    if (!Array.isArray(cookies)) cookies = [cookies];
    cookies = cookies.map(sanitize).filter(Boolean);

    if (!cookies.length) {
      return res.status(400).json({ error: "Chưa có cookie" });
    }

    const r = await axios.post(
      WORKER_URL + "/orders",
      { cookies },
      {
        headers: { "Content-Type": "application/json" },
        timeout: 20000,
        responseType: "arraybuffer",
      }
    );

    const html = Buffer.from(r.data).toString("utf8");

    if (typeof html === "string" && html.trim().startsWith("{")) {
      return res.status(502).json({ error: "Worker trả lỗi JSON", detail: html });
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
// Sheet1: A=MVD, B=ShopeeOrder, C=Name, D=Address, E=Product, F=COD
// ======================================================
// ======================================================
// 2) EXCEL MODE: /api/track?code=...
// Đọc theo HEADER (không phụ thuộc thứ tự cột)
// ======================================================
const EXCEL_PATH = path.join(__dirname, "data", "orders.xlsx");
let EXCEL_CACHE = null;
let EXCEL_MTIME = 0;

function cleanCell(v) {
  return String(v ?? "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();
}

function pick(row, keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== "") return row[k];
  }
  return "";
}

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

  // bỏ header (dòng 1)
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const A = ws[XLSX.utils.encode_cell({ r, c: 0 })]; // Mã đơn
    const B = ws[XLSX.utils.encode_cell({ r, c: 1 })]; // Tên
    const C = ws[XLSX.utils.encode_cell({ r, c: 2 })]; // Địa chỉ
    const D = ws[XLSX.utils.encode_cell({ r, c: 3 })]; // Sản phẩm
    const E = ws[XLSX.utils.encode_cell({ r, c: 4 })]; // COD
    const F = ws[XLSX.utils.encode_cell({ r, c: 5 })]; // Mã vận đơn
    const G = ws[XLSX.utils.encode_cell({ r, c: 6 })]; // Mã đơn Shopee

    const inputOrder = String(A?.v ?? "").trim().toUpperCase();     // mã đơn (cột A)
    const mvd = String(F?.v ?? "").trim().toUpperCase();           // mã vận đơn (cột F)
    const shopeeOrder = String(G?.v ?? "").trim().toUpperCase();   // mã đơn Shopee (cột G)

    // nếu cả 3 mã đều rỗng thì bỏ
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
    const found = list.find(x =>
      x.inputOrder === code || x.mvd === code || x.shopeeOrder === code
    );

    if (!found) return res.status(404).json({ error: "Không tìm thấy trong Excel" });
    return res.json({ order: found });
  } catch (e) {
    return res.status(500).json({ error: "Lỗi đọc Excel", detail: e.message });
  }
});


app.get("/api/track", (req, res) => {
  const code = String(req.query.code || "").trim().toUpperCase();
  if (!code) return res.status(400).json({ error: "Thiếu mã (code)" });

  try {
    const list = loadExcel();
    const found = list.find(x => x.mvd === code || x.shopeeOrder === code);
    if (!found) return res.status(404).json({ error: "Không tìm thấy trong Excel" });
    return res.json({ order: found });
  } catch (e) {
    return res.status(500).json({ error: "Lỗi đọc Excel", detail: e.message });
  }
});

// ======================================================
import multer from "multer";

const upload = multer({ storage: multer.memoryStorage() });

function requireAdmin(req, res, next) {
  const token = req.headers["x-admin-token"];
  if (!process.env.ADMIN_TOKEN) {
    return res.status(500).json({ error: "Missing ADMIN_TOKEN on server" });
  }
  if (token !== process.env.ADMIN_TOKEN) {
    return res.status(401).json({ error: "Unauthorized" });
  }
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

    // reset cache để đọc file mới ngay
    EXCEL_CACHE = null;
    EXCEL_MTIME = 0;

    return res.json({ ok: true, message: "Đã cập nhật Excel" });
  } catch (e) {
    return res.status(500).json({ error: "Upload lỗi", detail: e.message });
  }
});

app.listen(3000, () => {
  console.log("Server chạy: http://localhost:3000");
  console.log("Excel path:", EXCEL_PATH);
});
