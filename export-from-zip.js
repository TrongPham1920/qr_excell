const express = require("express");
const multer = require("multer");
const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const path = require("path");

const app = express();
const port = 3000;

const upload = multer({ storage: multer.memoryStorage() });

/**
 * Chuẩn hóa tên file:
 * - Bỏ extension
 * - Bỏ tiền tố ._
 * - Bỏ dấu '
 * - Trim khoảng trắng
 */
function normalizeFileName(fileName) {
  let name = path.parse(fileName).name;

  // Bỏ tiền tố rác macOS
  if (name.startsWith("._")) {
    name = name.replace("._", "");
  }

  // Bỏ dấu nháy '
  name = name.replace(/'/g, "");

  // Trim khoảng trắng
  name = name.trim();

  return name;
}

app.post("/upload-zip", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Không có file upload" });
    }

    const zip = new AdmZip(req.file.buffer);
    const zipEntries = zip.getEntries();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");

    worksheet.columns = [
      { header: "STT", key: "stt", width: 10 },
      { header: "SERIAL", key: "serial", width: 30 },
      { header: "IMAGE", key: "image", width: 25 },
    ];

    let stt = 1;
    let rowIndex = 2;

    for (const entry of zipEntries) {
      if (entry.isDirectory) continue;

      const fileName = path.basename(entry.entryName);

      // ❌ Bỏ file rác macOS
      if (fileName.startsWith("._") || fileName === ".DS_Store") {
        continue;
      }

      const ext = path.extname(fileName).toLowerCase();

      // Chỉ xử lý ảnh
      if (![".png", ".jpg", ".jpeg"].includes(ext)) {
        continue;
      }

      const serial = normalizeFileName(fileName);

      // Nếu tên rỗng thì bỏ qua
      if (!serial) continue;

      // Thêm row text
      worksheet.addRow({
        stt: stt++,
        serial: serial,
      });

      // Thêm ảnh vào Excel
      const imageId = workbook.addImage({
        buffer: entry.getData(),
        extension: ext.replace(".", ""),
      });

      worksheet.addImage(imageId, {
        tl: { col: 2, row: rowIndex - 1 },
        ext: { width: 100, height: 100 },
      });

      worksheet.getRow(rowIndex).height = 80;

      rowIndex++;
    }

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );

    res.setHeader("Content-Disposition", "attachment; filename=export.xlsx");

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(error);
    res.status(500).json({ message: "Lỗi server" });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
