const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const convertRouter = require('./routes/convert');

const app = express();
const port = process.env.PORT || 5001;

// Настройка CORS
app.use(cors({
  origin: ['http://176.124.219.69:3002', 'http://localhost:3002'],
  methods: ['GET', 'POST'],
  credentials: true
}));

app.use(express.json());

// Логирование всех запросов
app.use((req, res, next) => {
  console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
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

app.listen(port, '0.0.0.0', () => {
  console.log(`Server is running on port ${port}`);
});
