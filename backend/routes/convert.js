const express = require('express');
const router = express.Router();
const convertService = require('../services/convertService');

router.post('/', async (req, res) => {
  console.log('Получен POST запрос на /api/convert');
  try {
    if (!req.file) {
      console.log('Файл не был загружен');
      return res.status(400).send('Файл не был загружен');
    }

    console.log('Файл получен:', req.file);
    console.log('Начало конвертации файла');
    const discountPercentage = req.body.discountPercentage ? parseFloat(req.body.discountPercentage) : null;
    const makeShortVersion = req.body.makeShortVersion === 'true';
    const originalFileName = req.body.originalFileName;

    const buffer = await convertService.convertExcelToWord(req.file.path, discountPercentage, makeShortVersion, originalFileName);
    console.log('Конвертация завершена успешно');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.docx');
    res.send(buffer);
    console.log('Файл отправлен клиенту');
  } catch (error) {
    console.error('Ошибка при конвертации:', error);
    res.status(500).send(`Произошла ошибка при конвертации файла: ${error.message}`);
  }
});

module.exports = router;
