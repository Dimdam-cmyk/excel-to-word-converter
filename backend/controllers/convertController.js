const convertService = require('../services/convertService');

exports.convertExcelToWord = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Файл не загружен');
    }

    console.log('Начало конвертации файла:', req.file.path);

    const wordBuffer = await convertService.convertExcelToWord(req.file.path);

    console.log('Конвертация успешно завершена');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.docx');
    res.send(wordBuffer);
  } catch (error) {
    console.error('Ошибка при конвертации файла:', error);
    console.error('Стек ошибки:', error.stack);
    res.status(500).send(`Произошла ошибка при конвертации файла: ${error.message}\n\nСтек ошибки: ${error.stack}`);
  }
};
