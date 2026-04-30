const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;
const DATA_DIR = path.join(__dirname, "data");
const JSON_PATH = path.join(DATA_DIR, "records.json");
const SHEET_NAME = "Records";

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

function ensureDataFile() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  if (!fs.existsSync(JSON_PATH)) {
    fs.writeFileSync(JSON_PATH, JSON.stringify([], null, 2), "utf8");
  }
}

function readRecords() {
  ensureDataFile();
  const raw = fs.readFileSync(JSON_PATH, "utf8");
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeRecords(records) {
  ensureDataFile();
  fs.writeFileSync(JSON_PATH, JSON.stringify(records, null, 2), "utf8");
}

function normalizeRecord(item) {
  const amountRaw = item.amount ?? item["金额"] ?? item["Amount"];
  const amountNumber = Number(amountRaw);
  const no = item.no ?? item["No."] ?? item["No"] ?? "";
  const date = item.date ?? item["Date"] ?? "";
  const type = item.type ?? item["支出/收入"] ?? "";
  const category = item.category ?? item["类型"] ?? "";
  const note = item.note ?? item["其他备注"] ?? item["备注"] ?? "";

  return {
    id: String(item.id || Date.now() + Math.random()),
    no: String(no),
    date: String(date),
    type: type === "收入" ? "收入" : "支出",
    category: String(category),
    amount: Number.isNaN(amountNumber) ? 0 : amountNumber,
    note: String(note)
  };
}

function toExcelRows(records) {
  return records.map((item, index) => ({
    "No.": item.no || String(index + 1),
    Date: item.date || "",
    "支出/收入": item.type || "",
    类型: item.category || "",
    金额: Number(item.amount || 0),
    其他备注: item.note || ""
  }));
}

app.get("/api/records", (req, res) => {
  try {
    const records = readRecords();
    res.json(records);
  } catch (error) {
    res.status(500).json({ message: "读取数据失败", error: error.message });
  }
});

app.post("/api/records", (req, res) => {
  try {
    const { date, type, category, amount, note } = req.body;

    if (!date || !type || !category || amount === undefined || amount === null) {
      return res.status(400).json({ message: "缺少必要字段" });
    }

    const amountNumber = Number(amount);
    if (Number.isNaN(amountNumber)) {
      return res.status(400).json({ message: "金额必须是数字" });
    }

    const records = readRecords();
    const newRecord = normalizeRecord({
      id: Date.now().toString(),
      no: String(records.length + 1),
      date,
      type,
      category,
      amount: amountNumber,
      note: note || ""
    });

    records.push(newRecord);
    writeRecords(records);
    res.status(201).json(newRecord);
  } catch (error) {
    res.status(500).json({ message: "保存数据失败", error: error.message });
  }
});

app.delete("/api/records/:id", (req, res) => {
  try {
    const { id } = req.params;
    const records = readRecords();
    const nextRecords = records.filter((item) => String(item.id) !== String(id));

    if (nextRecords.length === records.length) {
      return res.status(404).json({ message: "未找到记录" });
    }

    for (let i = 0; i < nextRecords.length; i += 1) {
      nextRecords[i].no = String(i + 1);
    }
    writeRecords(nextRecords);
    res.json({ message: "删除成功" });
  } catch (error) {
    res.status(500).json({ message: "删除失败", error: error.message });
  }
});

app.post("/api/load-excel", (req, res) => {
  try {
    const { base64 } = req.body;
    if (!base64) {
      return res.status(400).json({ message: "缺少 Excel 文件内容" });
    }

    const buffer = Buffer.from(base64, "base64");
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return res.status(400).json({ message: "Excel 没有可读取的 sheet" });
    }

    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    const incomingRecords = rows.map(normalizeRecord);
    const existingRecords = readRecords();
    const records = existingRecords.concat(incomingRecords);
    for (let i = 0; i < records.length; i += 1) {
      records[i].no = String(i + 1);
    }
    writeRecords(records);

    res.json({ message: "Excel 导入成功（已追加）", count: incomingRecords.length });
  } catch (error) {
    res.status(500).json({ message: "Excel 导入失败", error: error.message });
  }
});

app.get("/api/save-excel", (req, res) => {
  try {
    const records = readRecords();
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(toExcelRows(records));
    XLSX.utils.book_append_sheet(workbook, worksheet, SHEET_NAME);
    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="ledger-${Date.now()}.xlsx"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: "Excel 导出失败", error: error.message });
  }
});

app.listen(PORT, () => {
  ensureDataFile();
  console.log(`Server running at http://localhost:${PORT}`);
});
