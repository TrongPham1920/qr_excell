const express = require("express");
const router = express.Router();
const multer = require("multer");
const qrController = require("../controllers/qrController");

const upload = multer({ storage: multer.memoryStorage() });

/**
 * @swagger
 * /qr/upload-zip:
 *   post:
 *     summary: Upload ZIP file chứa PNG QR
 *     tags: [QR]
 *     requestBody:
 *       required: true
 *       content:
 *         multipart/form-data:
 *           schema:
 *             type: object
 *             properties:
 *               file:
 *                 type: string
 *                 format: binary
 *     responses:
 *       200:
 *         description: Xuất Excel thành công
 *         content:
 *           application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *             schema:
 *               type: string
 *               format: binary
 *       400:
 *         description: File ZIP không hợp lệ
 *       500:
 *         description: Lỗi server
 */
router.post("/upload-zip", upload.single("file"), qrController.uploadZip);

/**
 * @swagger
 * /qr/upload-excel:
 *   post:
 *     summary: Upload Excel chứa serial và QR Image
 *     tags: [QR]
 *     requestBody:
 *       required: true
 *       content:
 *         multipart/form-data:
 *           schema:
 *             type: object
 *             properties:
 *               file:
 *                 type: string
 *                 format: binary
 *     responses:
 *       200:
 *         description: Xuất Excel mới thành công
 *         content:
 *           application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *             schema:
 *               type: string
 *               format: binary
 *       400:
 *         description: File Excel không hợp lệ
 *       500:
 *         description: Lỗi server
 */
router.post("/upload-excel", upload.single("file"), qrController.uploadExcel);

module.exports = router;
