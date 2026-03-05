const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const path = require("path");
const helper = require("./helper");

const processExcell = async(buffer) => {
    const zip = new AdmZip(buffer);
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

      const serial = helper.normalizeFileName(fileName);
      if (!serial) continue;

      const imageBuffer = entry.getData();

      // Đọc QR
      const lapValue = await helper.readQrFromPng(imageBuffer);

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
    return workbook;
}

module.exports = {
  processExcell,
};