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
    const includeVAT = req.body.includeVAT === 'true';
    const originalFileName = req.body.originalFileName;
    const rbtChecked = req.body.rbtChecked === 'true';
    const extraExpenses = req.body.extraExpenses ? JSON.parse(req.body.extraExpenses) : [];
    const rbtDiscount = req.body.rbtDiscount ? parseFloat(req.body.rbtDiscount) : null;
    const pricePerSqm = req.body.pricePerSqm ? parseFloat(req.body.pricePerSqm) : 14000;
    const subsystemPercentage = req.body.subsystemPercentage ? parseFloat(req.body.subsystemPercentage) : 15;

    const result = await convertService.convertExcelToWord(req.file.path, discountPercentage, makeShortVersion, originalFileName, includeVAT, rbtChecked, extraExpenses, rbtDiscount, pricePerSqm, subsystemPercentage);
    console.log('Конвертация завершена успешно');
    console.log('Тип результата:', typeof result);
    console.log('Результат содержит wordBuffer:', !!(result && result.wordBuffer));
    console.log('Результат содержит pdfBuffer:', !!(result && result.pdfBuffer));

    // Если результат содержит PDF для монтажа, создаем ZIP архив
    if (result && typeof result === 'object' && result.wordBuffer && result.pdfBuffer) {
      const JSZip = require('jszip');
      const zip = new JSZip();
      
      zip.file('commercial_offer.docx', result.wordBuffer);
      zip.file('mounting_offer.docx', result.pdfBuffer);
      
      const zipBuffer = await zip.generateAsync({ type: 'nodebuffer' });
      
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename=offers.zip');
      res.send(zipBuffer);
    } else {
      // Обычный режим - только Word документ
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', 'attachment; filename=converted.docx');
      res.send(result);
    }
    console.log('Файл отправлен клиенту');
  } catch (error) {
    console.error('Ошибка при конвертации:', error);
    res.status(500).send(`Произошла ошибка при конвертации файла: ${error.message}`);
  }
});

module.exports = router;
