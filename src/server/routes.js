const express = require('express');
const router = express.Router();

const qrRoute = require('../routes/qrRoute');

router.use('/qr', qrRoute);

module.exports = router;