import React, { useState } from 'react';
import { Container, Typography, Button, CircularProgress } from '@material-ui/core';
import { makeStyles } from '@material-ui/core/styles';
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
}));

function App() {
  const classes = useStyles();
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleFileChange = (selectedFile) => {
    setFile(selectedFile);
  };

  const handleConvert = async () => {
    if (!file) {
      alert('Пожалуйста, выберите файл Excel');
      return;
    }

    setLoading(true);

    try {
      const response = await convertExcelToWord(file);
      const blob = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'converted.docx';
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Ошибка при конвертации:', error);
      alert('Произошла ошибка при конвертации файла');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container className={classes.container}>
      <Typography variant="h4" gutterBottom>
        Конвертер Excel в Word
      </Typography>
      <FileUploader onFileChange={handleFileChange} />
      <Button
        variant="contained"
        color="primary"
        onClick={handleConvert}
        disabled={!file || loading}
        className={classes.button}
      >
        {loading ? <CircularProgress size={24} /> : 'Конвертировать'}
      </Button>
    </Container>
  );
}

export default App;
