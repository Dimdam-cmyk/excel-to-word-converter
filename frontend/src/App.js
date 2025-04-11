import React, { useState, useEffect } from 'react';
import { Container, Typography, Button, CircularProgress, Snackbar, Checkbox, TextField, Paper, Box, FormControlLabel, LinearProgress } from '@material-ui/core';
import { makeStyles } from '@material-ui/core/styles';
import MuiAlert from '@material-ui/lab/Alert';
import FileUploader from './components/FileUploader';
import { convertExcelToWord } from './services/api';

const useStyles = makeStyles((theme) => ({
  container: {
    marginTop: theme.spacing(4),
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    marginBottom: theme.spacing(4),
  },
  paper: {
    padding: theme.spacing(4),
    borderRadius: 16,
    boxShadow: '0 8px 30px rgba(0, 0, 0, 0.12)',
    width: '100%',
    maxWidth: 600,
    background: 'linear-gradient(to bottom, #ffffff, #f8f9fa)',
  },
  title: {
    marginBottom: theme.spacing(3),
    fontWeight: 600,
    color: '#2c3e50',
    textAlign: 'center',
  },
  button: {
    marginTop: theme.spacing(4),
    padding: '12px 24px',
    borderRadius: 8,
    fontWeight: 'bold',
    letterSpacing: 1,
    background: 'linear-gradient(45deg, #4568dc, #5e56dc)',
    '&:hover': {
      background: 'linear-gradient(45deg, #3456c8, #4744c4)',
      boxShadow: '0 8px 15px rgba(69, 104, 220, 0.3)',
    },
    transition: 'all 0.3s ease',
    textTransform: 'none',
  },
  checkboxGroup: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    marginTop: theme.spacing(3),
    padding: theme.spacing(2),
    backgroundColor: 'rgba(242, 245, 250, 0.6)',
    borderRadius: 8,
  },
  checkboxLabel: {
    marginLeft: -8,
  },
  checkboxItem: {
    margin: theme.spacing(0.5, 0),
  },
  discountInput: {
    marginTop: theme.spacing(2),
    width: '100%',
    '& .MuiOutlinedInput-root': {
      borderRadius: 8,
    },
  },
  logo: {
    width: '540px',
    marginBottom: theme.spacing(2),
    filter: 'drop-shadow(0 4px 6px rgba(0, 0, 0, 0.1))',
  },
  loader: {
    color: '#ffffff',
  },
  progressContainer: {
    width: '100%',
    marginTop: theme.spacing(3),
    padding: theme.spacing(2),
    backgroundColor: 'rgba(242, 245, 250, 0.6)',
    borderRadius: 8,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
  },
  progressText: {
    marginBottom: theme.spacing(1),
    color: '#2c3e50',
    fontWeight: 500,
  },
  progress: {
    width: '100%',
    height: 10,
    borderRadius: 5,
    '& .MuiLinearProgress-barColorPrimary': {
      background: 'linear-gradient(45deg, #4568dc, #5e56dc)',
    },
  },
  progressResult: {
    marginTop: theme.spacing(2),
    padding: theme.spacing(2),
    backgroundColor: 'rgba(69, 104, 220, 0.1)',
    borderRadius: 8,
    width: '100%',
    textAlign: 'center',
  },
}));

function Alert(props) {
  return <MuiAlert elevation={6} variant="filled" {...props} />;
}

function App() {
  const classes = useStyles();
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [applyDiscount, setApplyDiscount] = useState(false);
  const [discountPercentage, setDiscountPercentage] = useState('');
  const [makeShortVersion, setMakeShortVersion] = useState(false);
  const [includeVAT, setIncludeVAT] = useState(false);
  const [progressPercentage, setProgressPercentage] = useState(0);
  const [progressText, setProgressText] = useState('');
  const [showProgress, setShowProgress] = useState(false);
  const [progressComplete, setProgressComplete] = useState(false);
  const [selectedEmployee, setSelectedEmployee] = useState('');
  const [selectedCar, setSelectedCar] = useState('');

  const funnyTexts = [
    'Накидываем на скидку...',
    'Прикручиваем 30% за вредность для клиента...',
    'Добавляем коэффициент жадности...',
    'Закладываем на новую квартиру для кладовщицы...',
    'На дачу у Моря для Турсуна...',
    'Учитываем бонус на корпоративные вечеринки...',
    'Увеличиваем бюджет на печеньки в переговорной...',
    'Закладываем на элитный фитнес для бухгалтерии...',
    'Компенсация за стресс при общении с клиентами...',
    'Резерв на непредвиденные расходы директора...'
  ];

  const employees = [
    'Бурдейный Дмитрий',
    'Влад Гейинле',
    'Анастасия Пирумова',
    'Анастасия Пигина',
    'Екатерина Поплавская',
    'Виталий Ильин',
    'Дмитрий Чумичев',
    'Ксения Демме',
    'Анна Маркетолог',
    'Карина Фасилитатор'
  ];

  const cars = [
    'BMW X5',
    'Mercedes S-Class',
    'Audi Q7',
    'Toyota Land Cruiser',
    'Lexus LX570',
    'Porsche Cayenne',
    'Range Rover Sport',
    'Jaguar F-Pace',
    'Bentley Bentayga',
    'Tesla Model X',
    'Ferrari 488',
    'Lamborghini Urus',
    'Maserati Levante',
    'Rolls-Royce Cullinan',
    'Cadillac Escalade',
    'Infiniti QX80',
    'Volvo XC90',
    'Genesis GV80'
  ];

  const getRandomItem = (array) => {
    return array[Math.floor(Math.random() * array.length)];
  };

  const handleFileChange = (selectedFile) => {
    console.log('Файл выбран:', selectedFile.name);
    setFile(selectedFile);
  };

  useEffect(() => {
    let progressInterval;
    let textInterval;
    let isTextUpdateActive = true; // Флаг для контроля активности обновления текста
    
    // Инициализируем финальные значения
    const finalEmployee = getRandomItem(employees);
    const finalCar = getRandomItem(cars);
    
    if (showProgress) {
      // Установим начальное сообщение сразу
      setProgressText('Начинаем конвертацию...');
      
      // Обновление прогресса каждые 100ms для заполнения за 10 секунд (100 * 100ms = 10000ms)
      progressInterval = setInterval(() => {
        setProgressPercentage(prev => {
          if (prev < 100) return prev + 1;
          
          // При достижении 100%, останавливаем обновление текста
          isTextUpdateActive = false;
          clearInterval(textInterval);
          
          // Устанавливаем финальные значения для результирующего сообщения
          setSelectedEmployee(finalEmployee);
          setSelectedCar(finalCar);
          
          // Устанавливаем окончательные значения
          setProgressComplete(true);
          
          // Устанавливаем финальный текст
          console.log('Установка финального текста');
          setProgressText(`Готово! Файл сконвертирован.`);
          
          clearInterval(progressInterval);
          return 100;
        });
      }, 100);
      
      // Обновление текста каждые 2500ms (2,5 секунды)
      textInterval = setInterval(() => {
        // Только если прогресс еще не завершен
        if (isTextUpdateActive) {
          console.log('Обновляем текст прогресса'); // для дебага
          if (Math.random() > 0.6) {
            // 40% вероятность показать сообщение с автомобилем и сотрудником
            const employee = getRandomItem(employees);
            const car = getRandomItem(cars);
            setSelectedEmployee(employee);
            setSelectedCar(car);
            const newText = `На новую ${car} для ${employee}...`;
            console.log('Новый текст:', newText); // для дебага
            setProgressText(newText);
          } else {
            // 60% вероятность показать обычное шуточное сообщение
            const randomText = getRandomItem(funnyTexts);
            console.log('Шуточный текст:', randomText); // для дебага
            setProgressText(randomText);
          }
        }
      }, 2500);
    }
    
    return () => {
      clearInterval(progressInterval);
      clearInterval(textInterval);
    };
  }, [showProgress]);

  const handleConvert = async () => {
    if (!file) {
      setError('Пожалуйста, выберите файл Excel');
      return;
    }
    
    // Сбрасываем состояние перед началом новой конвертации
    setShowProgress(true);
    setProgressPercentage(0);
    setProgressComplete(false);
    setProgressText('Начинаем конвертацию...');
    setLoading(true);
    setError(null);

    // Запускаем конвертацию параллельно с прогресс баром
    try {
      console.log('Начало конвертации файла:', file.name);
      const response = await convertExcelToWord(file, applyDiscount ? discountPercentage : null, makeShortVersion, includeVAT);
      console.log('Ответ получен:', response);

      // После получения ответа просто ждем небольшую задержку, чтобы был виден прогресс
      // Прогресс-бар сам закончится через 6 секунд
      
      // Небольшая задержка для лучшего UX
      setTimeout(() => {
        // Скачиваем файл
        const blob = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'converted.docx';
        a.click();
        window.URL.revokeObjectURL(url);
        console.log('Файл успешно сконвертирован и скачан');
        
        // Выключаем индикатор загрузки через секунду после скачивания
        setTimeout(() => {
          setLoading(false);
        }, 1000);
      }, 1000);
      
    } catch (error) {
      console.error('Ошибка при конвертации:', error);
      setError(error.response?.data || 'Произошла ошибка при конвертации файла');
      setLoading(false);
    }
  };

  const handleCloseError = (event, reason) => {
    if (reason === 'clickaway') {
      return;
    }
    setError(null);
  };

  return (
    <Container className={classes.container}>
      <img src="/logo.png" alt="Logo" className={classes.logo} />
      
      <Paper className={classes.paper} elevation={0}>
        <Typography variant="h4" className={classes.title}>
          Конвертер Excel в Word
        </Typography>
        
        <FileUploader onFileChange={handleFileChange} />
        
        <Box className={classes.checkboxGroup}>
          <FormControlLabel
            className={classes.checkboxItem}
            control={
              <Checkbox
                checked={applyDiscount}
                onChange={(e) => setApplyDiscount(e.target.checked)}
                color="primary"
              />
            }
            label="Добавить скидку"
          />
          
          {applyDiscount && (
            <TextField
              className={classes.discountInput}
              label="Процент скидки"
              type="number"
              value={discountPercentage}
              onChange={(e) => setDiscountPercentage(e.target.value)}
              variant="outlined"
              size="small"
              placeholder="Например: 10"
            />
          )}
          
          <FormControlLabel
            className={classes.checkboxItem}
            control={
              <Checkbox
                checked={makeShortVersion}
                onChange={(e) => setMakeShortVersion(e.target.checked)}
                color="primary"
              />
            }
            label="Сделать сокращ. КП"
          />
          
          <FormControlLabel
            className={classes.checkboxItem}
            control={
              <Checkbox
                checked={includeVAT}
                onChange={(e) => setIncludeVAT(e.target.checked)}
                color="primary"
              />
            }
            label="С НДС"
          />
        </Box>
        
        {showProgress && (
          <Box className={classes.progressContainer}>
            <Typography variant="body1" className={classes.progressText}>
              {progressText || 'Начинаем конвертацию...'}
            </Typography>
            <LinearProgress 
              variant="determinate" 
              value={progressPercentage} 
              className={classes.progress}
            />
            {progressComplete && selectedEmployee && selectedCar && (
              <Box className={classes.progressResult}>
                <Typography variant="body1">
                  Если заказчик заключит сделку, то {selectedEmployee} сможет ездить на работу на {selectedCar}
                </Typography>
              </Box>
            )}
          </Box>
        )}
        
        <Button
          variant="contained"
          color="primary"
          fullWidth
          onClick={handleConvert}
          disabled={!file || loading}
          className={classes.button}
        >
          {loading ? <CircularProgress size={24} className={classes.loader} /> : 'Конвертировать'}
        </Button>
      </Paper>
      
      <Snackbar open={!!error} autoHideDuration={6000} onClose={handleCloseError}>
        <Alert onClose={handleCloseError} severity="error">
          {error}
        </Alert>
      </Snackbar>
    </Container>
  );
}

export default App;
