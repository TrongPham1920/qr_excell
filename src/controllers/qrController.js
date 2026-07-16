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

const valueToText = (value) => {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();

  if (typeof value === "object") {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => valueToText(part.text)).join("");
    }
    if (Object.prototype.hasOwnProperty.call(value, "result")) {
      return valueToText(value.result);
    }
    if (Object.prototype.hasOwnProperty.call(value, "text")) {
      return valueToText(value.text);
    }
    if (Object.prototype.hasOwnProperty.call(value, "hyperlink")) {
      return valueToText(value.hyperlink);
    }
  }

  return "";
};

const getCellText = (cell) => {
  if (!cell) return "";

  const valueText = valueToText(cell.value).trim();
  if (valueText) return valueText;

  if (typeof cell.text === "string" && cell.text !== "[object Object]") {
    return cell.text.trim();
  }

  return "";
};

const normalizeHeader = (value) =>
  valueToText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

const isSttHeader = (header) => {
  const compact = header.replace(/\s+/g, "");
  return compact === "stt" || compact === "no" || compact === "sothutu";
};

const isSerialHeader = (header) => {
  const compact = header.replace(/\s+/g, "");
  return compact === "serial" || compact === "soserial";
};

const isLpaHeader = (header) => {
  const compact = header.replace(/\s+/g, "");
  return compact === "lpa" || compact === "lap";
};

const isQrCodeHeader = (header) => {
  const compact = header.replace(/\s+/g, "");
  return compact === "qrcode" || compact === "qr";
};

const detectHeaderColumns = (row) => {
  const columns = {
    sttCol: -1,
    serialCol: -1,
    lpaCol: -1,
    qrCol: -1,
  };

  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const header = normalizeHeader(getCellText(cell));
    if (!header) return;

    if (columns.sttCol === -1 && isSttHeader(header)) columns.sttCol = colNumber;
    if (columns.serialCol === -1 && isSerialHeader(header)) columns.serialCol = colNumber;
    if (columns.lpaCol === -1 && isLpaHeader(header)) columns.lpaCol = colNumber;
    if (columns.qrCol === -1 && isQrCodeHeader(header)) columns.qrCol = colNumber;
  });

  return columns;
};

const findDataSheet = (workbook) => {
  let candidate = null;

  workbook.eachSheet((sheet) => {
    if (candidate && candidate.serialCol !== -1 && candidate.lpaCol !== -1) return;

    const header = detectHeaderColumns(sheet.getRow(1));
    const hasKnownColumn =
      header.serialCol !== -1 ||
      header.lpaCol !== -1 ||
      header.qrCol !== -1;

    if (hasKnownColumn) {
      candidate = {
        sheet,
        headerRowNumber: 1,
        ...header,
      };
    }
  });

  return candidate;
};

const getImageAnchor = (img) => {
  if (img.range && img.range.tl) {
    return {
      col: img.range.tl.nativeCol + 1,
      row: img.range.tl.nativeRow + 1,
    };
  }

  if (img.range && img.range.from) {
    return {
      col: img.range.from.col + 1,
      row: img.range.from.row + 1,
    };
  }

  return null;
};

const findImageMedia = (media, imageId) =>
  media.find(
    (item) =>
      item.index === imageId ||
      item.id === imageId ||
      String(item.index) === String(imageId) ||
      String(item.name) === String(imageId),
  );

const buildImageMap = (inputWorkbook, sheet, qrCol) => {
  const map = {};
  const media = inputWorkbook.model.media || [];
  const apiImages = typeof sheet.getImages === "function" ? sheet.getImages() : [];
  const modelImages = sheet.model._images || [];
  const images = [...apiImages, ...modelImages];

  images.forEach((img) => {
    const anchor = getImageAnchor(img);
    if (!anchor || anchor.col !== qrCol) return;
    const m = findImageMedia(media, img.imageId);

    if (m && m.buffer) map[anchor.row] = m.buffer;
  });

  return map;
};

const uniqueColumns = (...columns) =>
  [...new Set(columns.filter((col) => col !== -1))];

const uploadExcel = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Không có file upload" });
    }

    if (!req.file.buffer || req.file.buffer.length === 0) {
      return res.status(400).json({ message: "File upload trống" });
    }

    console.log("File:", req.file.originalname, "size:", req.file.buffer.length);

    let xlsxBuffer = req.file.buffer;
    const ext = path.extname(req.file.originalname).toLowerCase();
    if (ext === ".xls") {
      try {
        const fullWb = XLSX.read(req.file.buffer, { type: "buffer", cellDates: false });
        if (!fullWb.SheetNames || fullWb.SheetNames.length === 0) {
          return res.status(400).json({ message: "File .xls không có sheet dữ liệu" });
        }
        xlsxBuffer = XLSX.write(fullWb, { bookType: "xlsx", type: "buffer" });
      } catch (xlsErr) {
        return res.status(400).json({ message: "File .xls không hợp lệ hoặc bị hỏng" });
      }
    }

    const inputWorkbook = new ExcelJS.Workbook();
    await inputWorkbook.xlsx.load(xlsxBuffer);

    const dataSheet = findDataSheet(inputWorkbook);
    if (!dataSheet) {
      return res.status(400).json({
        message: 'Dòng 1 phải có header "Số Serial" hoặc "Serial", và có ít nhất một cột "LPA" hoặc "QR CODE". QR CODE có thể là text LPA hoặc ảnh QR.',
      });
    }

    const {
      sheet: inputSheet,
      headerRowNumber,
      serialCol,
      lpaCol,
      qrCol,
    } = dataSheet;

    if (serialCol === -1) {
      return res.status(400).json({
        message: 'Thiếu cột Serial. Dòng 1 phải có header "Số Serial" hoặc "Serial".',
      });
    }

    if (lpaCol === -1 && qrCol === -1) {
      return res.status(400).json({
        message: 'Thiếu cột LPA hoặc QR CODE. Dòng 1 cần có "LPA" để chứa text, hoặc "QR CODE" để chứa text LPA/ảnh QR.',
      });
    }

    const imageMaps = uniqueColumns(qrCol).map((col) =>
      buildImageMap(inputWorkbook, inputSheet, col),
    );
    const hasImages = imageMaps.some((imageMap) => Object.keys(imageMap).length > 0);
    const getCellImageBuffer = (rowNumber) => {
      for (const imageMap of imageMaps) {
        if (imageMap[rowNumber]) return imageMap[rowNumber];
      }

      return null;
    };

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
    const invalidRows = [];

    for (let i = headerRowNumber + 1; i <= totalRows; i++) {
      const row = inputSheet.getRow(i);

      const serial = getCellText(row.getCell(serialCol));
      let qrValue = lpaCol !== -1 ? getCellText(row.getCell(lpaCol)) : "";
      const qrCodeText = qrCol !== -1 ? getCellText(row.getCell(qrCol)) : "";

      const cellImageBuffer = hasImages ? getCellImageBuffer(i) : null;

      if (!qrValue && qrCodeText) {
        qrValue = qrCodeText;
      }

      if (cellImageBuffer && !qrValue) {
        const decoded = await format.readQrFromPng(cellImageBuffer);
        if (decoded) qrValue = decoded;
      }

      if (!serial && !qrValue && !cellImageBuffer) continue;
      if (!serial) {
        invalidRows.push(`dòng ${i}: thiếu Serial`);
        continue;
      }
      if (!qrValue) {
        invalidRows.push(`dòng ${i}: thiếu LPA/QR CODE hoặc không đọc được ảnh trong cột QR CODE`);
        continue;
      }

      outputSheet.addRow({ stt: stt++, serial, lpa: qrValue });

      let finalImageBuffer = null;

      if (cellImageBuffer) {
        finalImageBuffer = cellImageBuffer;
      } else if (qrValue) {
        try {
          finalImageBuffer = await QRCode.toBuffer(qrValue, {
            type: "png",
            errorCorrectionLevel: "M",
            margin: 1,
            width: 300,
          });
        } catch (qrErr) {
          console.warn(`Không tạo được QR ở dòng ${i}:`, qrErr.message);
        }
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

    if (invalidRows.length > 0) {
      return res.status(400).json({
        message: `File có dòng dữ liệu chưa đúng: ${invalidRows.slice(0, 5).join("; ")}${invalidRows.length > 5 ? "; ..." : ""}`,
      });
    }

    if (stt === 1) {
      return res.status(400).json({
        message: "Không tìm thấy dòng dữ liệu hợp lệ sau header. Mỗi dòng cần có Serial và LPA text, hoặc QR CODE dạng text/ảnh.",
      });
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
