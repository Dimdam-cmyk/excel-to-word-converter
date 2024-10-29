import React, { useState } from 'react';
import { Container, Typography, Button, CircularProgress, Snackbar, Checkbox, TextField } from '@material-ui/core';
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
  },
  button: {
    marginTop: theme.spacing(2),
  },
  checkbox: {
    marginTop: theme.spacing(2),
  },
  discountInput: {
    marginTop: theme.spacing(2),
    width: '100%',
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

  const handleFileChange = (selectedFile) => {
    console.log('Файл выбран:', selectedFile.name);
    setFile(selectedFile);
  };

  const handleConvert = async () => {
    if (!file) {
      setError('Пожалуйста, выберите файл Excel');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      console.log('Начало конвертации файла:', file.name);
      const response = await convertExcelToWord(file, applyDiscount ? discountPercentage : null, makeShortVersion);
      console.log('Ответ получен:', response);

      const blob = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'converted.docx';
      a.click();
      window.URL.revokeObjectURL(url);
      console.log('Файл успешно сконвертирован и скачан');
    } catch (error) {
      console.error('Ошибка при конвертации:', error);
      setError(error.response?.data || 'Произошла ошибка при конвертации файла');
    } finally {
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
      <Typography variant="h4" gutterBottom>
        Конвертер Excel в Word
      </Typography>
      <FileUploader onFileChange={handleFileChange} />
      <div className={classes.checkbox}>
        <Checkbox
          checked={applyDiscount}
          onChange={(e) => setApplyDiscount(e.target.checked)}
          color="primary"
        />
        <Typography component="span">Добавить скидку</Typography>
      </div>
      {applyDiscount && (
        <TextField
          className={classes.discountInput}
          label="Процент скидки"
          type="number"
          value={discountPercentage}
          onChange={(e) => setDiscountPercentage(e.target.value)}
          variant="outlined"
          size="small"
        />
      )}
      <div className={classes.checkbox}>
        <Checkbox
          checked={makeShortVersion}
          onChange={(e) => setMakeShortVersion(e.target.checked)}
          color="primary"
        />
        <Typography component="span">Сделать сокращ. КП</Typography>
      </div>
      <Button
        variant="contained"
        color="primary"
        onClick={handleConvert}
        disabled={!file || loading}
        className={classes.button}
      >
        {loading ? <CircularProgress size={24} /> : 'Конвертировать'}
      </Button>
      <Snackbar open={!!error} autoHideDuration={6000} onClose={handleCloseError}>
        <Alert onClose={handleCloseError} severity="error">
          {error}
        </Alert>
      </Snackbar>
    </Container>
  );
}

export default App;
