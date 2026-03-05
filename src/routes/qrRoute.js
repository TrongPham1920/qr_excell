const express = require("express");
const router = express.Router();
const multer = require("multer");
const qrController = require("../controllers/qrController");

const upload = multer({ storage: multer.memoryStorage() });

router.route("/upload-zip")
    .post(upload.single("file"), qrController.uploadZip)

module.exports = router;

