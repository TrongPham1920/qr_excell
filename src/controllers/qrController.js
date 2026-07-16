const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const path = require("path");
const format = require("../utils/format");
const QRCode = require("qrcode");
const XLSX = require("xlsx");

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

const buildImageMap = (inputWorkbook, qrCol) => {
  const map = {};
  const media = inputWorkbook.model.media || [];

  inputWorkbook.eachSheet((sheet) => {
    const images = sheet.model._images || [];
    images.forEach((img) => {
      if (!img.range || !img.range.from) return;
      const fromCol = img.range.from.col + 1;
      if (fromCol !== qrCol) return;
      const row = img.range.from.row + 1;
      const m = media.find(
        (item) =>
          item.id === img.imageId ||
          String(item.name) === String(img.imageId),
      );
      if (m && m.buffer) map[row] = m.buffer;
    });
  });

  return map;
};

const uploadExcel = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Không có file upload" });
    }

    let xlsxBuffer = req.file.buffer;
    const ext = path.extname(req.file.originalname).toLowerCase();
    if (ext === ".xls") {
      const wb = XLSX.read(req.file.buffer, { type: "buffer" });
      xlsxBuffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
    }

    const inputWorkbook = new ExcelJS.Workbook();
    await inputWorkbook.xlsx.load(xlsxBuffer);

    const inputSheet = inputWorkbook.getWorksheet(1);
    if (!inputSheet) {
      return res.status(400).json({
        message: "File Excel không có sheet dữ liệu",
      });
    }

    const headerRow = inputSheet.getRow(1);
    const normalize = (v) => String(v || "").trim().toLowerCase();

    let sttCol = -1;
    let serialCol = -1;
    let qrCol = -1;

    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const h = normalize(cell.value);
      if (h.includes("stt")) sttCol = colNumber;
      if (h.includes("serial")) serialCol = colNumber;
      if (["lpa", "lap", "qr code", "qr", "qr_code"].includes(h)) {
        qrCol = colNumber;
      }
    });

    if (serialCol === -1 || qrCol === -1) {
      return res.status(400).json({
        message: 'File phải có cột "Serial" và cột "QR CODE" hoặc "LPA"',
      });
    }

    const imageMap = buildImageMap(inputWorkbook, qrCol);
    const hasImages = Object.keys(imageMap).length > 0;

    const outputWorkbook = new ExcelJS.Workbook();
    const outputSheet = outputWorkbook.addWorksheet("Data");

    outputSheet.columns = [
      { header: "STT", key: "stt", width: 8 },
      { header: "Số Serial", key: "serial", width: 30 },
      { header: "LPA", key: "lpa", width: 80 },
      { header: "QR CODE", key: "image", width: 25 },
    ];

    let stt = 1;
    let outputRowIndex = 2;
    const totalRows = inputSheet.rowCount;

    for (let i = 2; i <= totalRows; i++) {
      const row = inputSheet.getRow(i);

      const serial = String(row.getCell(serialCol).value || "").trim();
      let qrValue = String(row.getCell(qrCol).value || "").trim();

      const cellImageBuffer = hasImages ? imageMap[i] || null : null;

      if (cellImageBuffer && !qrValue) {
        const decoded = await format.readQrFromPng(cellImageBuffer);
        if (decoded) qrValue = decoded;
      }

      if (!serial && !qrValue) continue;

      outputSheet.addRow({ stt: stt++, serial, lpa: qrValue });

      let finalImageBuffer = null;

      if (cellImageBuffer) {
        finalImageBuffer = cellImageBuffer;
      } else if (qrValue) {
        finalImageBuffer = await QRCode.toBuffer(qrValue, {
          type: "png",
          errorCorrectionLevel: "M",
          margin: 1,
          width: 300,
        });
      }

      if (finalImageBuffer) {
        const imageId = outputWorkbook.addImage({
          buffer: finalImageBuffer,
          extension: "png",
        });

        outputSheet.addImage(imageId, {
          tl: { col: 3, row: outputRowIndex - 1 },
          ext: { width: 160, height: 160 },
        });

        outputSheet.getRow(outputRowIndex).height = 120;
      }

      outputRowIndex++;
    }

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader("Content-Disposition", "attachment; filename=export.xlsx");

    await outputWorkbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(error);
    res.status(500).json({
      message: "Lỗi server",
      error: error.message,
    });
  }
};

module.exports = { uploadZip, uploadExcel };
