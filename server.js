// Load environment variables
require('dotenv').config();

// Core dependencies
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

// For Excel saving
const ExcelJS = require('exceljs');

// For voice file uploads
const multer = require('multer');

const app = express();
app.use(cors());
app.use(express.json());

// Serve frontend files
app.use(express.static(path.join(__dirname, 'public')));

// Serve raw voice files so we can play them in browser
app.use('/voices', express.static(path.join(__dirname, 'voices')));

// Load constants from .env
const DEVICE_API_KEY = process.env.DEVICE_API_KEY;
const PORT = process.env.PORT || 5000;

// =============================
// EXCEL SETUP
// =============================
const excelFile = path.join(__dirname, "attendance.xlsx");
let workbook = new ExcelJS.Workbook();

async function initExcel() {
  if (fs.existsSync(excelFile)) {
    await workbook.xlsx.readFile(excelFile);
  } else {
    workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Attendance");

    sheet.columns = [
      { header: "Timestamp", key: "timestamp", width: 25 },
      { header: "Device ID", key: "device_id", width: 20 },
      { header: "QR Token", key: "token", width: 20 },
      { header: "PIN", key: "pin", width: 10 },
      { header: "Method", key: "method", width: 15 }
    ];

    await workbook.xlsx.writeFile(excelFile);
  }
}

initExcel();

// =============================
// ATTENDANCE DEVICE API
// =============================
app.post("/api/device/event", async (req, res) => {
  const key = req.headers["x-device-key"];
  if (key !== DEVICE_API_KEY) {
    return res.status(401).json({ error: "invalid key" });
  }

  const event = req.body.events[0];
  console.log("Attendance event:", event);

  const sheet = workbook.getWorksheet("Attendance");

  sheet.addRow({
    timestamp: new Date().toISOString(),
    device_id: req.body.device_id,
    token: event.token_or_pin,
    pin: event.extras ? event.extras.entered_pin : "",
    method: event.method
  });

  await workbook.xlsx.writeFile(excelFile);

  res.json({
    results: [
      { ok: true, message: "Attendance saved to Excel" }
    ]
  });
});

// =============================
// VOICE RECORDING STORAGE SETUP
// =============================
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const dir = path.join(__dirname, 'voices');
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    cb(null, dir);
  },
  filename: function (req, file, cb) {
    const studentId = req.body.student_id || 'unknown';
    const type = req.body.type || 'check';
    const timestamp = Date.now();
    cb(null, `${studentId}_${type}_${timestamp}.webm`);
  }
});

const upload = multer({ storage: storage });

// =============================
// VOICE ENROLLMENT ROUTE
// =============================
app.post("/api/voice/enroll", upload.single('audio'), (req, res) => {
  console.log("Voice ENROLL for:", req.body.student_id, "file:", req.file.filename);
  res.json({
    ok: true,
    message: "Voice enrolled and saved",
    file: req.file.filename
  });
});

// =============================
// VOICE CHECK ROUTE
// =============================
app.post("/api/voice/check", upload.single('audio'), (req, res) => {
  console.log("Voice CHECK for:", req.body.student_id, "file:", req.file.filename);
  res.json({
    ok: true,
    message: "Voice sample saved for review",
    file: req.file.filename
  });
});

// =============================
// LIST ALL VOICE FILES (for review)
// =============================
app.get("/api/voices", (req, res) => {
  const dir = path.join(__dirname, 'voices');
  if (!fs.existsSync(dir)) {
    return res.json([]);
  }

  const files = fs.readdirSync(dir).filter(f => f.endsWith('.webm'));

  const data = files.map(f => {
    // filename format: studentId_type_timestamp.webm
    const parts = f.replace('.webm', '').split('_');
    const student_id = parts[0] || 'unknown';
    const type = parts[1] || 'unknown';
    const ts = parts[2] ? new Date(Number(parts[2])) : null;

    return {
      student_id,
      type,
      filename: f,
      url: `/voices/${f}`,
      timestamp: ts ? ts.toISOString() : null
    };
  });

  res.json(data);
});

// =============================
// ATTENDANCE DASHBOARD ROUTE
// =============================
app.get("/api/attendance", async (req, res) => {
  const sheet = workbook.getWorksheet("Attendance");

  const rows = sheet.getSheetValues()
    .slice(2)
    .map(r => ({
      timestamp: r[1],
      device_id: r[2],
      token: r[3],
      pin: r[4],
      method: r[5]
    }));

  res.json(rows);
});

// =============================
// START SERVER
// =============================
app.listen(PORT, () => {
  console.log("SERVER + FRONTEND + VOICE running on port", PORT);
});
