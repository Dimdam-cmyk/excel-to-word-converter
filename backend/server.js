const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const convertRouter = require('./routes/convert');

const app = express();
const port = process.env.PORT || 5001; // Изменили порт на 5001

app.use(cors());
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

app.use('/api/convert', convertRouter);

app.listen(port, () => {
  console.log(`Сервер запущен на порту ${port}`);
});
