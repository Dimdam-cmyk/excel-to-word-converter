const express = require('express');
const multer = require('multer');
const convertController = require('../controllers/convertController');

const router = express.Router();
const upload = multer({ dest: 'uploads/' });

router.post('/', upload.single('file'), convertController.convertExcelToWord);

module.exports = router;
