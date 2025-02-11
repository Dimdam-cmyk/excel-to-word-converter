import axios from 'axios';

const API_URL = process.env.REACT_APP_API_URL || 'http://176.124.219.69:5001/api';

export const convertExcelToWord = async (file) => {
  const formData = new FormData();
  formData.append('file', file);

  try {
    const response = await axios.post(`${API_URL}/convert`, formData, {
      headers: {
        'Content-Type': 'multipart/form-data',
      },
      responseType: 'arraybuffer',
      withCredentials: true
    });
    return response;
  } catch (error) {
    if (error.response) {
      const errorMessage = new TextDecoder().decode(error.response.data);
      throw new Error(errorMessage);
    }
    throw error;
  }
};
