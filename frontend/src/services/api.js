import axios from 'axios';

const API_URL = 'https://архио-коммерческое.рф/api';

export const convertExcelToWord = async (file, discountPercentage, makeShortVersion, includeVAT, rbtChecked, extraExpenses, rbtDiscount, pricePerSqm, subsystemPercentage) => {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('originalFileName', file.name);
  if (discountPercentage !== null) {
    formData.append('discountPercentage', discountPercentage);
  }
  formData.append('makeShortVersion', makeShortVersion);
  formData.append('includeVAT', includeVAT);
  formData.append('rbtChecked', rbtChecked);
  if (extraExpenses) {
    formData.append('extraExpenses', JSON.stringify(extraExpenses));
  }
  if (rbtDiscount !== null) {
    formData.append('rbtDiscount', rbtDiscount);
  }
  if (pricePerSqm !== null) {
    formData.append('pricePerSqm', pricePerSqm);
  }
  if (subsystemPercentage !== null) {
    formData.append('subsystemPercentage', subsystemPercentage);
  }

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
