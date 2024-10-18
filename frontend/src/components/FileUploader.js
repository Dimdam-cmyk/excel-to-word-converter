import React from 'react';
import { Button } from '@material-ui/core';
import { makeStyles } from '@material-ui/core/styles';

const useStyles = makeStyles((theme) => ({
  input: {
    display: 'none',
  },
}));

function FileUploader({ onFileChange }) {
  const classes = useStyles();

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    onFileChange(file);
  };

  return (
    <div>
      <input
        accept=".xlsx,.xls"
        className={classes.input}
        id="contained-button-file"
        type="file"
        onChange={handleFileChange}
      />
      <label htmlFor="contained-button-file">
        <Button variant="contained" color="default" component="span">
          Выбрать Excel файл
        </Button>
      </label>
    </div>
  );
}

export default FileUploader;
