const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const convertRouter = require('./routes/convert');

const app = express();
const port = process.env.PORT || 5001;

// Настройка CORS
app.use(cors({
  origin: [
    'http://176.124.219.69:3002', 
    'http://localhost:3002',
    'http://архио-коммерческое.рф',
    'https://архио-коммерческое.рф',  // Добавляем HTTPS версию
    'http://xn----7sbqaaopcpascfrir1d4b.xn--p1ai',
    'https://xn----7sbqaaopcpascfrir1d4b.xn--p1ai'  // Добавляем HTTPS версию punycode
  ],
  methods: ['GET', 'POST'],
  credentials: true
}));

// Логирование всех запросов
app.use((req, res, next) => {
  console.log('--------------------');
  console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
  console.log('Headers:', req.headers);
  if (req.file) console.log('File:', req.file);
  if (req.body) console.log('Body:', req.body);
  console.log('--------------------');
  next();
});

// Настройка multer для загрузки файлов
const upload = multer({ 
  dest: 'uploads/',
  fileFilter: (req, file, cb) => {
    if (file.mimetype === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      cb(null, true);
    } else {
      cb(new Error("Неподдерживаемый тип файла"), false);
    }
  },
  limits: {
    fileSize: 5 * 1024 * 1024, // Ограничение размера файла до 5 МБ
  },
});

// Обработчик для корневого маршрута
app.get('/', (req, res) => {
  res.send('Backend server is running');
});

// Использование маршрутов конвертации
app.use('/api/convert', upload.single('file'), convertRouter);

// Обработка ошибок
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).send('Что-то пошло не так!');
});

process.on('uncaughtException', (err) => {
    console.error('Необработанная ошибка:', err);
    // Логируем ошибку, но не завершаем процесс
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Необработанное отклонение промиса:', reason);
    // Логируем ошибку, но не завершаем процесс
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Server is running on port ${port}`);
});
