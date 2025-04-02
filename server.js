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
const bodyParser = require("body-parser");

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

    console.log("📷 Xử lý ảnh:", filePath);

    // Xử lý ảnh (grayscale, sharpen, resize để tối ưu nhận diện)
    await sharp(filePath)
      .grayscale()
      .sharpen()
      .resize(1600, 1400, { fit: "inside" })
      .toFormat("png")
      .toFile(processedFilePath);

    // Đọc ảnh thành buffer
    const imageBuffer = fs.readFileSync(processedFilePath);

    // Quét tất cả mã vạch có trong ảnh
    const results = await readBarcodesFromImageData(imageBuffer);
    const barcodes = results.map((result) => result.text); // Lưu danh sách mã vạch

    // Xóa ảnh gốc sau khi xử lý
    fs.unlinkSync(filePath);
    fs.unlinkSync(processedFilePath);

    if (barcodes.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "Không tìm thấy mã vạch trong ảnh!" });
    }

    console.log("✅ Mã vạch quét được:", barcodes);
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

    const filePath = path.join(exportDir, `output_${Date.now()}.xlsx`);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Barcode Data");

    // Tiêu đề cột
    worksheet.columns = [
      { header: "STT", key: "index", width: 10 },
      { header: "Mã Vạch", key: "barcode", width: 30 },
    ];

    // Thêm dữ liệu vào Excel
    data.forEach((barcode, index) => {
      worksheet.addRow({ index: index + 1, barcode });
    });

    // Xuất file Excel
    await workbook.xlsx.writeFile(filePath);

    res.json({ success: true, file: filePath });
  } catch (error) {
    console.error("❌ Lỗi xuất Excel:", error);
    res.status(500).json({ success: false, message: "Lỗi server!" });
  }
});

/**
 * 📌 API: Nhập dữ liệu vào Google Sheets
 */
const CREDENTIALS = JSON.parse(process.env.GOOGLE_CLOUD_CREDENTIALS); // Đảm bảo tệp JSON đúng

const auth = new google.auth.GoogleAuth({
  credentials: CREDENTIALS,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

app.post("/upload-google-sheet", async (req, res) => {
  try {
    const { sheetUrl, barcodes } = req.body;

    // ✅ Kiểm tra sheetUrl hợp lệ
    const match = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      return res
        .status(400)
        .json({ success: false, message: "❌ URL Google Sheet không hợp lệ!" });
    }
    const sheetId = match[1];
    console.log("📌 Sheet ID:", sheetId);

    // ✅ Kiểm tra barcodes hợp lệ
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
    console.log("📌 Barcodes gửi:", barcodes);
    //console.log("📌 Barcodes:", JSON.stringify(barcodes, null, 2));

    // ✅ Ghi dữ liệu vào Google Sheets
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

// Chạy server
app.listen(PORT, () => {
  console.log(`🚀 Server chạy trên http://localhost:${PORT}`);
});
