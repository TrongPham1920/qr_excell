const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const path = require("path");
const format = require("../utils/format");

const uploadZip = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Không có file upload" });
    } 

    const zip = new AdmZip(req.file.buffer);
    const zipEntries = zip.getEntries();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");

    worksheet.columns = [
      { header: "STT", key: "stt", width: 8 },
      { header: "SERIAL", key: "serial", width: 30 },
      { header: "LAP", key: "lap", width: 150 },
      { header: "IMAGE", key: "image", width: 25 },
    ];

    let stt = 1;
    let rowIndex = 2;

    for (const entry of zipEntries) {
      if (entry.isDirectory) continue;

      const fileName = path.basename(entry.entryName);

      if (fileName.startsWith("._") || fileName === ".DS_Store") continue;

      const ext = path.extname(fileName).toLowerCase();

      // 👉 Chỉ xử lý PNG cho ổn định
      if (ext !== ".png") continue;

      const serial = format.normalizeFileName(fileName);
      if (!serial) continue;

      const imageBuffer = entry.getData();

      // Đọc QR
      const lapValue = await format.readQrFromPng(imageBuffer);

      // Add text row
      worksheet.addRow({
        stt: stt++,
        serial: serial,
        lap: lapValue || "Không đọc được QR",
      });

      // Add image
      const imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: "png",
      });

      worksheet.addImage(imageId, {
        tl: { col: 3, row: rowIndex - 1 }, // IMAGE là cột thứ 4
        ext: { width: 160, height: 160 },
      });

      worksheet.getRow(rowIndex).height = 120;

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

};
module.exports = { uploadZip };