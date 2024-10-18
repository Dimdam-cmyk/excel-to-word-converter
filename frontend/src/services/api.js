import axios from 'axios';

const API_URL = 'http://localhost:5001/api'; // Изменили порт на 5001

export const convertExcelToWord = async (file) => {
  const formData = new FormData();
  formData.append('file', file);

  return axios.post(`${API_URL}/convert`, formData, {
    responseType: 'arraybuffer',
    headers: {
      'Content-Type': 'multipart/form-data',
    },
  });
};
