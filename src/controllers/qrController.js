const excell = require("../utils/excell");

const uploadZip = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Không có file upload" });
    } 

    const workbook = await excell.processExcell(req.file.buffer);

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