import React, { useState } from 'react';
import { Box, Button, Typography, Paper } from '@material-ui/core';
import { makeStyles } from '@material-ui/core/styles';
import CloudUploadIcon from '@material-ui/icons/CloudUpload';
import InsertDriveFileIcon from '@material-ui/icons/InsertDriveFile';

const useStyles = makeStyles((theme) => ({
  input: {
    display: 'none',
  },
  dropzone: {
    border: '2px dashed #9fc3e7',
    borderRadius: 12,
    padding: theme.spacing(4),
    textAlign: 'center',
    backgroundColor: 'rgba(242, 245, 250, 0.6)',
    transition: 'all 0.3s ease',
    cursor: 'pointer',
    marginBottom: theme.spacing(2),
    '&:hover': {
      borderColor: '#4568dc',
      backgroundColor: 'rgba(242, 245, 250, 0.9)',
    },
  },
  dragActive: {
    borderColor: '#4568dc',
    backgroundColor: 'rgba(69, 104, 220, 0.1)',
  },
  icon: {
    fontSize: 48,
    color: '#4568dc',
    marginBottom: theme.spacing(2),
  },
  fileInfo: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(1.5),
    borderRadius: 8,
    backgroundColor: 'rgba(242, 245, 250, 0.9)',
    marginTop: theme.spacing(2),
  },
  fileIcon: {
    color: '#4b6cb7',
    marginRight: theme.spacing(1),
  },
  fileName: {
    fontWeight: 500,
    color: '#2c3e50',
    flexGrow: 1,
  },
  uploadButton: {
    marginTop: theme.spacing(2),
    borderRadius: 8,
    textTransform: 'none',
    fontWeight: 500,
  }
}));

function FileUploader({ onFileChange }) {
  const classes = useStyles();
  const [isDragging, setIsDragging] = useState(false);
  const [selectedFile, setSelectedFile] = useState(null);

  const handleDragOver = (event) => {
    event.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (event) => {
    event.preventDefault();
    setIsDragging(false);
    
    const file = event.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      setSelectedFile(file);
      onFileChange(file);
    }
  };

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    setSelectedFile(file);
    onFileChange(file);
  };

  return (
    <Box width="100%">
      <input
        accept=".xlsx,.xls"
        className={classes.input}
        id="contained-button-file"
        type="file"
        onChange={handleFileChange}
      />
      <label 
        htmlFor="contained-button-file" 
        className={`${classes.dropzone} ${isDragging ? classes.dragActive : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <CloudUploadIcon className={classes.icon} />
        <Typography variant="body1" gutterBottom>
          Перетащите файл Excel сюда или нажмите для выбора
        </Typography>
        <Typography variant="body2" color="textSecondary">
          Поддерживаются файлы .xlsx и .xls
        </Typography>
      </label>

      {selectedFile && (
        <Paper elevation={0} className={classes.fileInfo}>
          <InsertDriveFileIcon className={classes.fileIcon} />
          <Typography className={classes.fileName} noWrap>
            {selectedFile.name}
          </Typography>
          <Typography variant="body2" color="textSecondary">
            {(selectedFile.size / 1024).toFixed(1)} КБ
          </Typography>
        </Paper>
      )}
    </Box>
  );
}

export default FileUploader;
