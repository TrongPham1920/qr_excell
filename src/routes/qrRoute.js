const express = require("express");
const router = express.Router();
const multer = require("multer");
const qrController = require("../controllers/qrController");

const upload = multer({ storage: multer.memoryStorage() });

/**
 * @swagger
 * /qr/upload-zip:
 *   post:
 *    summary: Upload a ZIP file containing PNG images with QR codes
 *    tags: [QR]
 *    requestBody:
 *     required: true
 *     content:
 *      multipart/form-data:
 *       schema:
 *        type: object
 *        properties:
 *         file:
 *          type: string
 *          format: binary
 *     description: Upload a ZIP file containing PNG images with QR codes.
 *    responses:
 *     200:
 *      description: Successfully processed the ZIP file and generated the Excel file.
 *      content:
 *        application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *          schema:
 *            type: string
 *            format: binary
 *     400:
 *      description: Bad request. Please upload a valid ZIP file.
 *     500:
 *      description: Internal server error. An error occurred while processing the ZIP file.
 */

router.route("/upload-zip")
    .post(upload.single("file"), qrController.uploadZip)

module.exports = router;

