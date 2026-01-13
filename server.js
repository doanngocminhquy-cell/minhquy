import express from "express";
import axios from "axios";
import * as cheerio from "cheerio";
import cors from "cors";
import path from "path";
import fs from "fs";
import XLSX from "xlsx";
import multer from "multer";
import { fileURLToPath } from "url";

const app = express();
app.use(cors());
app.use(express.json());

// ===== Path setup =====
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// static
app.use(express.static(path.join(__dirname, "public")));

// ===== Worker =====
const WORKER_URL = "https://1.doanngocminhquy.workers.dev";

// ===== Excel path =====
const EXCEL_PATH = path.join(__dirname, "data", "orders.xlsx");

// ===== cache =====
let EXCEL_CACHE = null;
let EXCEL_MTIME = 0;

// ===== helpers =====
function sanitize(s) {
  return String(s || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "")
    .trim();
}

// =================================================
// 1️⃣ SHOPEE COOKIE → Worker
// =================================================
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
        responseType: "arraybuffer"
      }
    );

    const html = Buffer.from(r.data).toString("utf8");

    const $ = cheerio.load(html);
    const orders = [];

    $("table tbody tr").each((_, tr) => {
      const tds = $(tr).find("td");
      if (tds.length < 9) return;

      orders.push({
        stt: $(tds[0]).text().trim(),
        productImg: $(tds[1]).find("img").attr("src") || null,
        cod: $(tds[2]).text().trim(),
        mvd: $(tds[3]).text().trim(),
        status: $(tds[4]).text().trim(),
        receiver: $(tds[5]).text().trim(),
        receiverPhone: $(tds[6]).text().trim(),
        address: $(tds[7]).attr("title") || $(tds[7]).text().trim(),
        shipperPhone: $(tds[8]).text().trim(),
      });
    });

    res.json({ count: orders.length, orders });
  } catch (e) {
    res.status(500).json({ error: "Worker lỗi", detail: e.message });
  }
});

// =================================================
// 2️⃣ EXCEL LOAD
// =================================================
function loadExcel() {
  const stat = fs.statSync(EXCEL_PATH);
  if (EXCEL_CACHE && EXCEL_MTIME === stat.mtimeMs) return EXCEL_CACHE;

  const buf = fs.readFileSync(EXCEL_PATH);
  const wb = XLSX.read(buf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws["!ref"]);

  const rows = [];

  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const code = String(ws["A" + (r + 1)]?.v || "").trim().toUpperCase();
    const name = String(ws["B" + (r + 1)]?.v || "").trim();
    const address = String(ws["C" + (r + 1)]?.v || "").trim();
    const product = String(ws["D" + (r + 1)]?.v || "").trim();
    const cod = String(ws["E" + (r + 1)]?.v || "").trim();
    const mvd = String(ws["F" + (r + 1)]?.v || "").trim();
    const shopee = String(ws["G" + (r + 1)]?.v || "").trim();

    if (!code && !mvd && !shopee) continue;

    rows.push({ code, name, address, product, cod, mvd, shopee });
  }

  EXCEL_CACHE = rows;
  EXCEL_MTIME = stat.mtimeMs;
  return rows;
}

// =================================================
// 3️⃣ TRA MÃ
// =================================================
app.get("/api/track", (req, res) => {
  const q = String(req.query.code || "").trim().toUpperCase();
  if (!q) return res.status(400).json({ error: "Thiếu mã" });

  try {
    const list = loadExcel();
    const found = list.find(x => x.code === q || x.mvd === q || x.shopee === q);

    if (!found) return res.status(404).json({ error: "Không tìm thấy" });

    res.json({ order: found });
  } catch (e) {
    res.status(500).json({ error: "Excel lỗi", detail: e.message });
  }
});

// =================================================
// 4️⃣ UPLOAD EXCEL (ADMIN)
// =================================================
const upload = multer({ storage: multer.memoryStorage() });

app.post("/api/upload-excel", upload.single("file"), (req, res) => {
  if (req.headers["x-admin-token"] !== process.env.ADMIN_TOKEN) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  if (!req.file) return res.status(400).json({ error: "No file" });

  const dir = path.join(__dirname, "data");
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  fs.writeFileSync(EXCEL_PATH, req.file.buffer);

  EXCEL_CACHE = null;
  EXCEL_MTIME = 0;

  res.json({ ok: true });
});

// =================================================
app.listen(3000, () => {
  console.log("Server running at http://localhost:3000");
});
