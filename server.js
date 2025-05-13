require("dotenv").config();
const express = require("express");
const multer = require("multer");
const sharp = require("sharp");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const cors = require("cors");
const { readBarcodesFromImageData } = require("zxing-wasm");
const { GoogleSpreadsheet } = require("google-spreadsheet");

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

// 📌 Tạo thư mục nếu chưa tồn tại
const ensureDirectoryExists = (dir) => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
};

const uploadDir = path.join(__dirname, "uploads");
const processedDir = path.join(__dirname, "processed_uploads");
const exportDir = path.join(__dirname, "exports");

ensureDirectoryExists(uploadDir);
ensureDirectoryExists(processedDir);
ensureDirectoryExists(exportDir);

// 📌 Cấu hình lưu ảnh tải lên
const storage = multer.diskStorage({
  destination: uploadDir,
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname));
  },
});
const upload = multer({ storage });

/**
 * 📌 API: Upload ảnh & quét nhiều mã vạch
 */
app.post("/upload", upload.single("image"), async (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, message: "Vui lòng tải lên một hình ảnh!" });
    }

    const filePath = req.file.path;
    const processedFilePath = path.join(processedDir, `${Date.now()}.png`);

    await sharp(filePath)
      .grayscale()
      .sharpen()
      .resize(1600, 1400, { fit: "inside" })
      .toFormat("png")
      .toFile(processedFilePath);

    const imageBuffer = fs.readFileSync(processedFilePath);

    const results = await readBarcodesFromImageData(imageBuffer);
    const barcodes = results.map((result) => result.text);

    fs.unlinkSync(filePath);
    fs.unlinkSync(processedFilePath);

    if (barcodes.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "Không tìm thấy mã vạch trong ảnh!" });
    }

    res.json({ success: true, barcodes });
  } catch (error) {
    console.error("❌ Lỗi quét mã vạch:", error);
    res.status(500).json({ success: false, message: "Lỗi server!" });
  }
});

/**
 * 📌 API: Xuất danh sách mã vạch ra file Excel
 */
app.post("/export-excel", async (req, res) => {
  try {
    const { data } = req.body;

    if (!data || !Array.isArray(data)) {
      return res
        .status(400)
        .json({ success: false, message: "Dữ liệu không hợp lệ!" });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Barcode Data");

    worksheet.columns = [
      { header: "STT", key: "index", width: 10 },
      { header: "Mã Vạch", key: "barcode", width: 30 },
    ];

    data.forEach((barcode, index) => {
      worksheet.addRow({ index: index + 1, barcode });
    });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename=barcode_data_${Date.now()}.xlsx`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("❌ Lỗi xuất Excel:", error);
    res.status(500).json({ success: false, message: "Lỗi server!" });
  }
});

/**
 * 📌 Khởi tạo Google Sheets & Server
 */
const startServer = async () => {
  try {
    const CREDENTIALS = JSON.parse(process.env.GOOGLE_CLOUD_CREDENTIALS);
    CREDENTIALS.private_key = CREDENTIALS.private_key.replace(/\\n/g, "\n");

    const auth = new google.auth.GoogleAuth({
      credentials: CREDENTIALS,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });

    const sheets = google.sheets({ version: "v4", auth });

    app.post("/upload-google-sheet", async (req, res) => {
      try {
        const { sheetUrl, barcodes } = req.body;

        const match = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) {
          return res
            .status(400)
            .json({
              success: false,
              message: "❌ URL Google Sheet không hợp lệ!",
            });
        }

        const sheetId = match[1];

        if (
          !Array.isArray(barcodes) ||
          barcodes.length === 0 ||
          !barcodes.every((b) => typeof b === "string")
        ) {
          return res.status(400).json({
            success: false,
            message: "❌ Dữ liệu phải là một mảng chuỗi hợp lệ!",
          });
        }

        await sheets.spreadsheets.values.append({
          spreadsheetId: sheetId,
          range: "Sheet1!A:A",
          valueInputOption: "RAW",
          insertDataOption: "INSERT_ROWS",
          requestBody: { values: barcodes.map((code) => [code]) },
        });

        res.json({
          success: true,
          message: "✅ Đã ghi dữ liệu vào Google Sheets!",
        });
      } catch (error) {
        console.error("❌ Google Sheets API Error:", error.message);
        res.status(500).json({
          success: false,
          message: "Lỗi server khi ghi Google Sheets!",
          error: error.message,
        });
      }
    });

    app.listen(PORT, () => {
      console.log(`🚀 Server chạy trên http://localhost:${PORT}`);
    });
  } catch (err) {
    console.error("❌ Lỗi khi khởi tạo Google Sheets:", err.message);
  }
};

startServer();
