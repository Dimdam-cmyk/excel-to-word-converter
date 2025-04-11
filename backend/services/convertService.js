const ExcelJS = require('exceljs');
const docx = require('docx');
const fs = require('fs');
const path = require('path');

function convertMillimetersToTwip(mm) {
  return Math.round(mm * 56.7);
}

function convertMillimetersToPixels(mm) {
  return Math.round(mm * 3.78); // 1 мм = 3.78 пикселей при 96 DPI
}

// Функция для расчета пропорционального масштабирования
function calculateProportionalDimensions(targetWidthMm, originalAspectRatio) {
  const targetWidthPx = convertMillimetersToPixels(targetWidthMm);
  const targetHeightPx = Math.round(targetWidthPx / originalAspectRatio);
  
  return { width: targetWidthPx, height: targetHeightPx };
}

// Функция преобразования числа в текстовое представление прописью
function numberToWordsRubles(num) {
  const units = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'];
  const teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать'];
  const tens = ['', 'десять', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
  const hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];
  const thousands = ['', 'тысяча', 'тысячи', 'тысяч'];
  const millions = ['', 'миллион', 'миллиона', 'миллионов'];
  
  if (num === 0) return 'ноль рублей';
  
  let words = '';
  
  // Обработка миллионов
  if (num >= 1000000) {
    const millions_num = Math.floor(num / 1000000);
    num = num % 1000000;
    
    if (millions_num === 1) {
      words += 'один миллион ';
    } else if (millions_num >= 2 && millions_num <= 4) {
      words += convertLessThanThousand(millions_num) + ' миллиона ';
    } else {
      words += convertLessThanThousand(millions_num) + ' миллионов ';
    }
  }
  
  // Обработка тысяч
  if (num >= 1000) {
    const thousands_num = Math.floor(num / 1000);
    num = num % 1000;
    
    if (thousands_num === 1) {
      words += 'одна тысяча ';
    } else if (thousands_num === 2) {
      words += 'две тысячи ';
    } else if (thousands_num >= 3 && thousands_num <= 4) {
      words += convertLessThanThousand(thousands_num) + ' тысячи ';
    } else {
      words += convertLessThanThousand(thousands_num) + ' тысяч ';
    }
  }
  
  // Обработка сотен, десятков и единиц
  if (num > 0) {
    words += convertLessThanThousand(num);
  }
  
  // Добавление "рублей" с правильным окончанием
  const lastDigit = num % 10;
  const lastTwoDigits = num % 100;
  
  if (lastTwoDigits >= 11 && lastTwoDigits <= 19) {
    words += ' рублей';
  } else if (lastDigit === 1) {
    words += ' рубль';
  } else if (lastDigit >= 2 && lastDigit <= 4) {
    words += ' рубля';
  } else {
    words += ' рублей';
  }
  
  return words.trim();
  
  // Вспомогательная функция для преобразования числа меньше 1000
  function convertLessThanThousand(n) {
    let result = '';
    
    // Сотни
    if (n >= 100) {
      result += hundreds[Math.floor(n / 100)] + ' ';
      n = n % 100;
    }
    
    // Десятки и единицы
    if (n >= 10 && n <= 19) {
      result += teens[n - 10];
    } else {
      if (n >= 20) {
        result += tens[Math.floor(n / 10)] + ' ';
        n = n % 10;
      }
      
      if (n > 0) {
        result += units[n];
      }
    }
    
    return result.trim();
  }
}

exports.convertExcelToWord = async (filePath, discountPercentage, makeShortVersion, originalFileName, includeVAT = false) => {
  console.log('=== Начало конвертации ===');
  console.log('Параметры:');
  console.log('- filePath:', filePath);
  console.log('- discountPercentage:', discountPercentage);
  console.log('- makeShortVersion:', makeShortVersion);
  console.log('- originalFileName:', originalFileName);
  console.log('- includeVAT:', includeVAT);

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    console.log('Excel файл успешно прочитан');
    console.log('Количество листов:', workbook.worksheets.length);
    console.log('Имена листов:', workbook.worksheets.map(ws => ws.name));

    const doc = new docx.Document({
      styles: {
        paragraphStyles: [
          {
            id: "totalRowStyle",
            name: "Total Row Style",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            run: {
              size: 22, // 11 пунктов = 22 half-points
              bold: true,
            },
            paragraph: {
              alignment: docx.AlignmentType.CENTER,
            },
          },
          {
            id: "italicStyle",
            name: "Italic Style",
            basedOn: "Normal",
            run: {
              italics: true,
            },
          },
        ],
      },
      sections: []
    });

    console.log('Создан пустой документ Word');

    const children = [];

    // Добавляем первое изображение (шапку)
    const headerImagePath = path.join(__dirname, '../assets/header.png');
    if (fs.existsSync(headerImagePath)) {
      try {
        // Получаем буфер изображения
        const headerImageBuffer = fs.readFileSync(headerImagePath);
        
        // Используем фиксированное соотношение сторон для изображения шапки
        // Увеличиваем соотношение сторон, чтобы сделать изображение менее высоким
        const originalAspectRatio = 6.5; // Увеличиваем соотношение ширина:высота
        
        // Целевая ширина по ширине таблицы
        const targetWidthMm = 250;
        
        // Рассчитываем высоту, сохраняя пропорции
        // Высота будет примерно 250/6.5 = 38.5 мм (около 4 см)
        const dimensions = calculateProportionalDimensions(targetWidthMm, originalAspectRatio);
        
        children.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: headerImageBuffer,
                transformation: {
                  width: dimensions.width,
                  height: dimensions.height,
                },
              }),
            ],
            spacing: { after: 300, before: 0 },
            alignment: docx.AlignmentType.CENTER, // Центрирование изображения
          })
        );
        console.log('Изображение шапки добавлено с размерами:', dimensions);
      } catch (error) {
        console.error('Ошибка при добавлении изображения шапки:', error);
      }
    } else {
      console.log('Файл изображения шапки не найден:', headerImagePath);
    }
    // Получаем имя файла без расширения и первые 10 символов
    const fileId = originalFileName ? path.parse(originalFileName).name : 'unknown';

    // Добавляем заголовок с идентификатором файла
    children.push(
      new docx.Paragraph({
        text: `Коммерческое предложение на поставку изделий из полимербетона ARHIO по проекту ${fileId}`,
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 0 },
        style: "Heading1"
      })
    );

    // Добавляем таблицу с данными
    const worksheet = workbook.getWorksheet(1);
    let tableRows = [];
    let totalSum = 0;

    console.log('Начало обработки строк Excel');

    // Группируем строки по значению в столбце "Наименование на фасаде"
    let groupedRows = [];
    let currentGroup = [];
    let currentName = '';

    // Добавляем программный заголовок таблицы
    const headerRow = new docx.TableRow({
      children: makeShortVersion ? [
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true })], 
          alignment: docx.AlignmentType.CENTER, 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" } 
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
      ] : [
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true })], 
          alignment: docx.AlignmentType.CENTER, 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" } // Добавляем серую заливку для заголовка
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Номенклатура', bold: true })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Цена, руб.', bold: true })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Площадь развёртки, м2', bold: true })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
      ],
    });
    tableRows.push(headerRow);

    // Начинаем чтение с 5-й строки
    for (let i = 5; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      const name = getCellValue(row.getCell('A'));
      
      if (name === currentName) {
        currentGroup.push(row);
      } else {
        if (currentGroup.length > 0) {
          groupedRows.push(currentGroup);
        }
        currentName = name;
        currentGroup = [row];
      }
      
      // Проверяем, достигли ли мы конца данных
      if (name && name.toString().includes('Итого стоимость производства составляет')) {
        break;
      }
    }

    // Добавляем последнюю группу, если она есть
    if (currentGroup.length > 0) {
      groupedRows.push(currentGroup);
    }

    // Создаем строки таблицы Word с объединенными ячейками
    let isEvenRow = false; // Флаг для чередования цвета фона

    if (groupedRows && Array.isArray(groupedRows)) {
      if (makeShortVersion) {
        // Код для короткой версии
        groupedRows.forEach(group => {
          const firstRow = group[0];
          const name = getCellValue(firstRow.getCell('A'));
          
          // Пропускаем строку "Итого стоимость работ составляет"
          if (name.includes('Итого стоимость производства составляет')) {
            return;
          }
          
          // Пропускаем строки с заголовками таблицы, которые случайно попадают в данные
          if (name === 'Наименование на фасаде' || name.toLowerCase().includes('наименование на фасаде')) {
            return;
          }
          
          let blockSum = 0;
          
          group.forEach(row => {
            blockSum += parseFloat(getCellValue(row.getCell('I'))) || 0;
          });

          const tableRow = new docx.TableRow({
            children: [
              new docx.TableCell({ children: [new docx.Paragraph({ text: name, alignment: docx.AlignmentType.LEFT })], verticalAlign: docx.VerticalAlign.CENTER }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(blockSum), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
            ],
          });
          tableRows.push(tableRow);
          
          totalSum += blockSum;
        });
      } else {
        // Код для полной версии
        groupedRows.forEach((group, groupIndex) => {
          const firstRow = group[0];
          const name = getCellValue(firstRow.getCell('A'));
          
          // Пропускаем строку "Итого стоимость работ составляет"
          if (name.includes('Итого стоимость производства составляет')) {
            return;
          }
          
          // Пропускаем строки с заголовками таблицы, которые случайно попадают в данные
          if (name === 'Наименование на фасаде' || name.toLowerCase().includes('наименование на фасаде')) {
            return;
          }
          
          let blockSum = 0;
          
          group.forEach((row, index) => {
            isEvenRow = !isEvenRow;
            const shading = isEvenRow ? { fill: "F2F2F2" } : undefined;

            const tableRow = new docx.TableRow({
              children: index === 0 ? 
                [
                  new docx.TableCell({
                    children: [new docx.Paragraph({ 
                      text: name,
                      alignment: docx.AlignmentType.CENTER,
                      style: "italicStyle"
                    })],
                    rowSpan: group.length,
                    verticalAlign: docx.VerticalAlign.CENTER,
                    width: { size: 2000, type: docx.WidthType.DXA },
                    properties: {
                      cantSplit: true,
                    }
                  }),
                  createCustomCell(getCellValue(row.getCell('B')), docx.AlignmentType.LEFT, shading, 1800),
                  createCustomCell(getCellValue(row.getCell('G')), docx.AlignmentType.CENTER, shading, 800),
                  createCustomCell(formatNumber(getCellValue(row.getCell('H'))), docx.AlignmentType.CENTER, shading, 1000),
                  createCustomCell(formatNumber(getCellValue(row.getCell('I'))), docx.AlignmentType.CENTER, shading, 1000),
                  createCustomCell(formatNumber(getCellValue(row.getCell('J')), 2), docx.AlignmentType.CENTER, shading, 1000),
                ] :
                [
                  createCustomCell(getCellValue(row.getCell('B')), docx.AlignmentType.LEFT, shading, 1800),
                  createCustomCell(getCellValue(row.getCell('G')), docx.AlignmentType.CENTER, shading, 800),
                  createCustomCell(formatNumber(getCellValue(row.getCell('H'))), docx.AlignmentType.CENTER, shading, 1000),
                  createCustomCell(formatNumber(getCellValue(row.getCell('I'))), docx.AlignmentType.CENTER, shading, 1000),
                  createCustomCell(formatNumber(getCellValue(row.getCell('J')), 2), docx.AlignmentType.CENTER, shading, 1000),
                ],
            });
            
            tableRows.push(tableRow);
            
            blockSum += parseFloat(getCellValue(row.getCell('I'))) || 0;
          });
          
          // Создаем итоговую строку с текстом "Итого:"
          const itogoParagraph = new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: "Итого:",
                bold: true,
                alignment: docx.AlignmentType.RIGHT
              })
            ],
            alignment: docx.AlignmentType.RIGHT
          });
          
          // Создаем пустую строку с жирным шрифтом для пропуска
          const emptyBoldParagraph = new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: "",
                bold: true
              })
            ]
          });
          
          // Создаем итоговую сумму с жирным шрифтом
          const totalSumParagraph = new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: formatNumber(blockSum),
                bold: true
              })
            ],
            alignment: docx.AlignmentType.CENTER
          });
          
          // Создаем итоговую площадь с жирным шрифтом
          const blockTotalArea = getBlockTotalArea(group);
          const totalAreaParagraph = new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: formatNumber(blockTotalArea, 2),
                bold: true
              })
            ],
            alignment: docx.AlignmentType.CENTER
          });
          
          // Добавляем итоговую строку для блока
          const blockTotalRow = new docx.TableRow({
            children: [
              new docx.TableCell({ 
                children: [itogoParagraph], 
                columnSpan: 3
              }),
              new docx.TableCell({ 
                children: [emptyBoldParagraph], 
                verticalAlign: docx.VerticalAlign.CENTER 
              }),
              new docx.TableCell({ 
                children: [totalSumParagraph], 
                verticalAlign: docx.VerticalAlign.CENTER 
              }),
              new docx.TableCell({ 
                children: [totalAreaParagraph], 
                verticalAlign: docx.VerticalAlign.CENTER 
              }),
            ],
            shading: { fill: "D9D9D9" },
          });
          tableRows.push(blockTotalRow);
          
          totalSum += blockSum;
        });
      }
    } else {
      console.log('groupedRows не определен или не является массивом');
    }

    console.log(`Обработано ${tableRows.length} строк`);

    // Получаем общую площадь из Excel файла
    let totalArea = 0;
    
    // Ищем строку с "Итого стоимость производства составляет" и берем значение площади из нее
    for (let i = 5; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      const cellA = getCellValue(row.getCell('A'));
      
      // Находим строку с итоговым значением
      if (cellA && cellA.toString().includes('Итого стоимость производства')) {
        // Берем значение из столбца J (площадь развертки)
        totalArea = parseFloat(getCellValue(row.getCell('J'))) || 0;
        console.log('Найдено итоговое значение площади развертки:', totalArea);
        break;
      }
      
      // Альтернативный подход - ищем строку, где в столбце A содержится "ИТОГО по проекту:"
      if (cellA && cellA.toString().includes('ИТОГО по проекту')) {
        totalArea = parseFloat(getCellValue(row.getCell('J'))) || 0;
        console.log('Найдено итоговое значение площади развертки (ИТОГО по проекту):', totalArea);
        break;
      }
    }
    
    // Если не нашли итоговое значение в таблице, используем площадь строки 37
    if (totalArea === 0) {
      try {
        const totalRow = worksheet.getRow(37); // Строка с "Площадь развертки изделий составляет, кв.м."
        if (totalRow) {
          const cellC = getCellValue(totalRow.getCell('C'));
          if (cellC && !isNaN(parseFloat(cellC))) {
            totalArea = parseFloat(cellC);
            console.log('Найдено значение площади развертки в строке 37:', totalArea);
          }
        }
      } catch (error) {
        console.log('Ошибка при поиске значения в строке 37:', error);
      }
    }
    
    // Если все равно не нашли, используем сумму площади всех строк
    if (totalArea === 0) {
      console.log('Не удалось найти итоговое значение площади, вычисляем сумму...');
      const processedRows = new Set();
      
      groupedRows.forEach(group => {
        group.forEach(row => {
          const rowId = `${getCellValue(row.getCell('B'))}-${getCellValue(row.getCell('G'))}`;
          if (!processedRows.has(rowId)) {
            processedRows.add(rowId);
            const area = parseFloat(getCellValue(row.getCell('J'))) || 0;
            totalArea += area;
          }
        });
      });
    }

    console.log('Итоговая площадь развертки:', totalArea);

    // Создаем параграфы для итоговой строки
    const totalProjectParagraph = new docx.Paragraph({
      children: [
        new docx.TextRun({
          text: 'ИТОГО по проекту:',
          bold: true
        })
      ],
      alignment: docx.AlignmentType.LEFT
    });
    
    const emptyBoldTotal = new docx.Paragraph({
      children: [
        new docx.TextRun({
          text: '',
          bold: true
        })
      ]
    });
    
    const totalSumParagraph = new docx.Paragraph({
      children: [
        new docx.TextRun({
          text: formatNumber(totalSum),
          bold: true
        })
      ],
      alignment: docx.AlignmentType.CENTER
    });
    
    const totalAreaParagraph = new docx.Paragraph({
      children: [
        new docx.TextRun({
          text: formatNumber(totalArea, 2),
          bold: true
        })
      ],
      alignment: docx.AlignmentType.CENTER
    });

    // Добавляем итоговую строку
    const totalSumRow = new docx.TableRow({
      children: makeShortVersion ? [
        new docx.TableCell({ 
          children: [totalProjectParagraph],
          verticalAlign: docx.VerticalAlign.CENTER,
        }),
        new docx.TableCell({ 
          children: [totalSumParagraph],
          verticalAlign: docx.VerticalAlign.CENTER 
        }),
      ] : [
        new docx.TableCell({ 
          children: [totalProjectParagraph], 
          columnSpan: 3, 
          verticalAlign: docx.VerticalAlign.CENTER,
        }),
        new docx.TableCell({ 
          children: [emptyBoldTotal], 
          verticalAlign: docx.VerticalAlign.CENTER 
        }),
        new docx.TableCell({ 
          children: [totalSumParagraph], 
          verticalAlign: docx.VerticalAlign.CENTER 
        }),
        new docx.TableCell({ 
          children: [totalAreaParagraph], 
          verticalAlign: docx.VerticalAlign.CENTER 
        }),
      ],
    });

    // Применяем заливку к ячейкам итоговой строки
    if (totalSumRow.children) {
      totalSumRow.children.forEach(cell => {
        if (cell) {
          cell.shading = { fill: "FFE699" }; // Золотистая заливка для итоговой строки
        }
      });
    }

    tableRows.push(totalSumRow);

    // Добавляем таблицу
    const firstTable = new docx.Table({
      rows: tableRows,
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      },
    });

    // Настройка ширины столбцов для первой таблицы
    if (tableRows.length > 0 && tableRows[0].children) {
      if (makeShortVersion) {
        // Настройка для короткой версии - только две колонки
        tableRows[0].children[0].width = { size: 4000, type: docx.WidthType.DXA }; // Наименование на фасаде - широкий
        tableRows[0].children[1].width = { size: 2000, type: docx.WidthType.DXA }; // Сумма - средний
      } else {
        // Настройка для полной версии
        // Наименование на фасаде - широкий
        tableRows[0].children[0].width = { size: 2000, type: docx.WidthType.DXA };
        
        // Номенклатура - средний
        tableRows[0].children[1].width = { size: 1800, type: docx.WidthType.DXA };
        
        // Кол-во изделий - узкий
        tableRows[0].children[2].width = { size: 800, type: docx.WidthType.DXA };
        
        // Цена - средний
        tableRows[0].children[3].width = { size: 1000, type: docx.WidthType.DXA };
        
        // Сумма - средний
        tableRows[0].children[4].width = { size: 1000, type: docx.WidthType.DXA };
        
        // Площадь развёртки - средний
        if (tableRows[0].children.length > 5) {
          tableRows[0].children[5].width = { size: 1000, type: docx.WidthType.DXA };
        }
      }
    }

    children.push(firstTable);

    // Добавляем новую таблицу "Стоимость форм и заливки"
    children.push(
      new docx.Paragraph({
        text: "Стоимость форм и заливки",
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 300 },
        style: "Heading1"
      })
    );

    // Попытка найти лист по имени
    let izdeliyaWorksheet = workbook.getWorksheet('Изделия');
    
    // Если лист не найден по имени, попробуем найти второй лист
    if (!izdeliyaWorksheet && workbook.worksheets.length > 1) {
      console.log('Лист "Изделия" не найден по имени, пробуем использовать второй лист');
      izdeliyaWorksheet = workbook.worksheets[1]; // Берем второй лист
    }
    
    if (izdeliyaWorksheet) {
      console.log('Найден лист для таблицы "Стоимость форм и заливки"');
      console.log('Имя найденного листа:', izdeliyaWorksheet.name);
      let tableRows = [];
      let isEvenRow = false;
      let totalH = 0, totalI = 0, totalK = 0, totalL = 0, totalN = 0;

      // Добавляем заголовок таблицы
      const headerRow = new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ text: '№ п/п', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Номенклатура', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Площадь развертки изделия, м2', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Площадь развертки общая, м2', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Масса общая, кг', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Кол-во форм, шт.', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость формы за м², руб.', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость форм для изделий, руб.', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость заливки за м², руб.', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость за единицу, руб.', bold: true })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
        ],
        tableHeader: true,
      });
      tableRows.push(headerRow);

      console.log('Начало обработки строк листа для "Стоимость форм и заливки"');
      console.log('Количество строк в листе:', izdeliyaWorksheet.rowCount);
      
      // Проверим первые несколько строк для отладки
      console.log('Данные первых строк:');
      for (let i = 1; i <= Math.min(5, izdeliyaWorksheet.rowCount); i++) {
        const row = izdeliyaWorksheet.getRow(i);
        console.log(`Строка ${i}: A=${getCellValue(row.getCell('A'))}, B=${getCellValue(row.getCell('B'))}, G=${getCellValue(row.getCell('G'))}, H=${getCellValue(row.getCell('H'))}`);
      }
      
      let processedRows = 0;
      for (let i = 2; i <= izdeliyaWorksheet.rowCount; i++) {
        const row = izdeliyaWorksheet.getRow(i);
        const cellA = getCellValue(row.getCell('A'));
        if (!cellA) {
          console.log(`Достигнут конец данных на строке ${i}`);
          break;
        }

        console.log(`Обработка строки ${i}`);
        console.log(`Данные: A=${cellA}, B=${getCellValue(row.getCell('B'))}, G=${getCellValue(row.getCell('G'))}, M=${getCellValue(row.getCell('M'))}, N=${getCellValue(row.getCell('N'))}, O=${getCellValue(row.getCell('O'))}, V=${getCellValue(row.getCell('V'))}`);
        
        isEvenRow = !isEvenRow;
        const shading = isEvenRow ? { fill: "F2F2F2" } : undefined;

        // Обновляем итоговые значения
        totalH += parseFloat(getCellValue(row.getCell('H'))) || 0;
        totalI += parseFloat(getCellValue(row.getCell('I'))) || 0;
        totalK += parseFloat(getCellValue(row.getCell('K'))) || 0;
        totalL += parseFloat(getCellValue(row.getCell('L'))) || 0;
        totalN += parseFloat(getCellValue(row.getCell('N'))) || 0;

        // Применяем скидку к значениям в столбцах M, N, O и V
        const originalM = parseFloat(getCellValue(row.getCell('M'))) || 0;
        const originalN = parseFloat(getCellValue(row.getCell('N'))) || 0;
        const originalO = parseFloat(getCellValue(row.getCell('O'))) || 0;
        const originalV = parseFloat(getCellValue(row.getCell('V'))) || 0;
        
        const discountedM = discountPercentage ? originalM * (1 - discountPercentage / 100) : originalM;
        const discountedN = discountPercentage ? originalN * (1 - discountPercentage / 100) : originalN;
        const discountedO = discountPercentage ? originalO * (1 - discountPercentage / 100) : originalO;
        const discountedV = discountPercentage ? originalV * (1 - discountPercentage / 100) : originalV;

        // Создаем ячейки с правильным форматированием
        const nomenklaturaCell = new docx.TableCell({ 
          children: [
            new docx.Paragraph({ 
              text: getCellValue(row.getCell('B')), 
              alignment: docx.AlignmentType.LEFT 
            })
          ], 
          verticalAlign: docx.VerticalAlign.CENTER, 
          shading,
          width: { size: 2500, type: docx.WidthType.DXA },
        });
        
        // Предотвращаем разрывы строк для номенклатуры
        nomenklaturaCell.properties = {
          cantSplit: true,
          noWrap: true,
        };
        
        const tableRow = new docx.TableRow({
          children: [
            new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('A')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            nomenklaturaCell,
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('G')), 3), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('H')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('I')), 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('K')), 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('L')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedM, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading: { fill: "FFE699" } }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedN, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedO, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading: { fill: "FFE699" } }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedV, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
          ],
        });
        tableRows.push(tableRow);
        processedRows++;
      }

      console.log(`Обработано ${processedRows} строк для второй таблицы`);

      // Добавляем итоговую строку
      const totalTableRow = new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ text: '', bold: true })],
            columnSpan: 3,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalH, 0),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalI, 2),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalK, 2),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalL, 0),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: '',
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalN * (1 - discountPercentage / 100), 2),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: '',
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: '',
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "FFE699" }
          }),
        ],
        height: {
          value: 400,
          rule: docx.HeightRule.ATLEAST,
        },
      });

      // Больше не нужно применять жирный шрифт для ячеек, так как уже прописано выше
      tableRows.push(totalTableRow);

      console.log('Создание таблицы');
      const secondTable = new docx.Table({
        rows: tableRows,
        width: {
          size: 100,
          type: docx.WidthType.PERCENTAGE,
        },
        // Настраиваем ширину для каждого столбца
        columnWidths: [400, 2500, 900, 700, 800, 700, 700, 900, 1000, 1000, 900],
      });

      // Настройка ширины столбцов через ячейки в заголовке
      if (tableRows.length > 0 && tableRows[0].children) {
        // № п/п - узкий
        tableRows[0].children[0].width = { size: 400, type: docx.WidthType.DXA };
        
        // Номенклатура - самый широкий
        tableRows[0].children[1].width = { size: 2500, type: docx.WidthType.DXA };
        
        // Площадь развертки изделия, м2
        tableRows[0].children[2].width = { size: 900, type: docx.WidthType.DXA };
        
        // Кол-во изделий, шт.
        tableRows[0].children[3].width = { size: 700, type: docx.WidthType.DXA };
        
        // Остальные столбцы - средней ширины
        for (let i = 4; i < tableRows[0].children.length; i++) {
          let width = 700;
          // Более широкие столбцы для стоимостей
          if (i >= 7) width = 900;
          tableRows[0].children[i].width = { size: width, type: docx.WidthType.DXA };
        }
      }

      console.log(`Размер таблицы: ${tableRows.length} строк (включая заголовок)`);
      
      if (tableRows.length > 1) {
        children.push(secondTable);
        console.log('Таблица "Стоимость форм и заливки" успешно добавлена');
      } else {
        console.log('Таблица "Стоимость форм и заливки" не содержит данных, пропускаем');
        
        // Добавляем сообщение о том, что таблица не содержит данных
        children.push(
          new docx.Paragraph({
            text: "Данные о стоимости форм и заливки отсутствуют",
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 300, before: 300 }
          })
        );
      }
    } else {
      console.log('Лист "Изделия" не найден');
    }

    console.log('Таблица "Стоимость форм и заливки" добавлена в документ Word');

    // Добавляем итоговую сумму
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Итого стоимость производства составляет ${formatNumber(totalSum)} руб. (${numberToWordsRubles(Math.round(totalSum))}${includeVAT ? ' включая НДС 20%' : ' без НДС'})`,
            bold: true,
          }),
        ],
        spacing: { before: 400, after: 200 },
      })
    );

    // Добавляем сумму НДС если чекбокс активирован
    if (includeVAT) {
      // Рассчитываем НДС (20% от общей суммы)
      const vatAmount = totalSum * 0.2 / 1.2; // НДС составляет 20% от суммы без НДС
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `В том числе НДС 20%: ${formatNumber(vatAmount)} руб.`,
              bold: true,
            }),
          ],
          spacing: { before: 200, after: 400 },
        })
      );
    }

    // Добавляем информацию о скидке, если она применяется
    let discountedTotal = totalSum;
    if (discountPercentage) {
      const discountAmount = totalSum * (discountPercentage / 100);
      discountedTotal = totalSum - discountAmount;

      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `Цена со скидкой ${discountPercentage}%: ${formatNumber(discountedTotal)} руб. (${numberToWordsRubles(Math.round(discountedTotal))}${includeVAT ? ' включая НДС 20%' : ' без НДС'})`,
              bold: true,
            }),
          ],
          spacing: { before: 200, after: 200 },
        })
      );

      // Добавляем сумму НДС для цены со скидкой, если чекбокс активирован
      if (includeVAT) {
        // Рассчитываем НДС (20% от суммы со скидкой)
        const vatAmountWithDiscount = discountedTotal * 0.2 / 1.2; // НДС составляет 20% от суммы без НДС
        children.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `В том числе НДС 20%: ${formatNumber(vatAmountWithDiscount)} руб.`,
                bold: true,
              }),
            ],
            spacing: { before: 200, after: 200 },
          })
        );
      }

      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `Скидка составила: ${formatNumber(discountAmount)} руб.`,
              bold: true,
            }),
          ],
          spacing: { before: 200, after: 400 },
        })
      );
    }

    // Добавляем дополнительную информацию из листа "Комплекты"
    const komplektyWorksheet = workbook.getWorksheet('Комплекты');
    if (komplektyWorksheet) {
      const additionalInfo = [
        { label: 'Цена 1 кв.м. развертки, руб.', column: 'C' },
        { label: 'Цена 1 кв.м. проекции, руб.', column: 'C' },
        { label: 'Площадь развертки изделий составляет, кв.м.', column: 'C' }
      ];

      for (const info of additionalInfo) {
        let foundRow = komplektyWorksheet.getRows(1, komplektyWorksheet.rowCount).find(row => row.getCell('A').value === info.label);
        if (foundRow) {
          let value = parseFloat(getCellValue(foundRow.getCell(info.column)));
          
          // Применяем скидку к ценам за кв.м., если скидка указана
          if (discountPercentage && (info.label.includes('Цена 1 кв.м.'))) {
            value = value * (1 - discountPercentage / 100);
          }
          
          children.push(new docx.Paragraph({ 
            text: `${info.label} ${formatNumberRounded(value)}${info.label.includes('Цена') ? ' (со скидкой)' : ''}`,
            spacing: { before: 200, after: 200 }
          }));
        }
      }
    }

    // Добавляем информацию о файле
    children.push(
      new docx.Paragraph({
        text: `Файл: ${path.basename(filePath)}`,
        spacing: { before: 400, after: 400 },
      })
    );

    // Добавляем изображения футера
    for (const imageName of ['footer1.png', 'footer2.png']) {
      const imagePath = path.join(__dirname, `../assets/${imageName}`);
      if (fs.existsSync(imagePath)) {
        try {
          const footerImageBuffer = fs.readFileSync(imagePath);
          children.push(
            new docx.Paragraph({
              children: [
                new docx.ImageRun({
                  data: footerImageBuffer,
                  transformation: {
                    width: convertMillimetersToPixels(250), // Устанавливаем ширину, равную ширине таблицы
                    height: convertMillimetersToPixels(150), // Уменьшаем высоту для лучшего соотношения
                  },
                }),
              ],
              alignment: docx.AlignmentType.CENTER,
            })
          );
          console.log(`Изображение ${imageName} добавлено`);
        } catch (error) {
          console.error(`Ошибка при добавлении изображения ${imageName}:`, error);
        }
      } else {
        console.log(`Файл изображения ${imageName} не найден:`, imagePath);
      }
    }

    // Получаем текущую дату
    const currentDate = new Date().toLocaleDateString('ru-RU');

    // Добавляем одну секцию со всем содержимым и колонтитулами
    doc.addSection({
      properties: {
        page: {
          size: {
            width: convertMillimetersToTwip(297),
            height: convertMillimetersToTwip(210),
          },
          orientation: docx.PageOrientation.LANDSCAPE,
          margins: {
            top: convertMillimetersToTwip(10),
            right: convertMillimetersToTwip(10),
            bottom: convertMillimetersToTwip(10),
            left: convertMillimetersToTwip(10),
          },
        },
      },
      headers: {
        default: new docx.Header({
          children: [
            new docx.Paragraph({
              text: `Дата составления предложения ${currentDate}`,
              alignment: docx.AlignmentType.RIGHT,
            }),
          ],
        }),
      },
      footers: {
        default: new docx.Footer({
          children: [
            new docx.Paragraph({
              text: "Предложение действительно 25 дней. Расчет является предварительным. Для окончательного расчета требуется проектирование.",
              alignment: docx.AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children: children,
    });

    console.log('Таблица, изображения и колонтитулы добавлены в документ Word');

    console.log('Начало создания буфера документа Word');
    const buffer = await docx.Packer.toBuffer(doc);
    console.log('Буфер документа Word создан');

    fs.unlinkSync(filePath);
    console.log('Временный файл Excel удален');
    return buffer;
  } catch (error) {
    console.error('Ошибка в процессе конвертации:', error);
    console.error('Стек вызовов:', error.stack);
    
    // Добавляем больше информации об ошибке
    if (error instanceof Error) {
      console.error('Имя ошибки:', error.name);
      console.error('Сообщение ошибки:', error.message);
    }
    
    if (error.code) {
      console.error('Код ошибки:', error.code);
    }
    
    if (error.syscall) {
      console.error('Системный вызов:', error.syscall);
    }
    if (error && error.name === 'ExcelJS.Error') {
      console.error('Ошибка ExcelJS:', error.message);
    } else if (error && error.name === 'docx.Error') {
      console.error('Ошибка docx:', error.message);
    }
    
    // Проверяем существование файла
    if (!fs.existsSync(filePath)) {
      console.error('Файл не найден:', filePath);
    } else {
      console.log('Размер файла:', fs.statSync(filePath).size, 'байт');
    }
    
    throw new Error(`Ошибка при конвертации файла: ${error.message}`);
  } finally {
    // Убедимся, что временный файл удаляется даже при возникновении ошибки
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
        console.log('Временный файл Excel удален');
      }
    } catch (unlinkError) {
      console.error('Ошибка при удалении временного файла:', unlinkError);
    }
  }
};

function getCellValue(cell) {
  if (!cell) return '';
  if (cell.formula) {
    return cell.result?.toString() || '';
  }
  return cell.value?.toString() || '';
}

function formatNumber(value, decimalPlaces = 2) {
  if (!value) return '';
  const num = parseFloat(value);
  if (isNaN(num)) return value;
  return num.toFixed(decimalPlaces).replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

function formatNumberRounded(value) {
  if (!value) return '';
  const num = parseFloat(value);
  if (isNaN(num)) return value;
  return Math.round(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

function calculateTotalArea(group) {
  let totalArea = 0;
  group.forEach(row => {
    const area = parseFloat(getCellValue(row.getCell('J'))) || 0;
    totalArea += area;
  });
  return totalArea;
}

function getBlockTotalArea(group) {
  let totalArea = 0;
  group.forEach(row => {
    const area = parseFloat(getCellValue(row.getCell('J'))) || 0;
    totalArea += area;
  });
  return totalArea;
}

function createCustomCell(text, alignment, shading, width) {
  return new docx.TableCell({
    children: [new docx.Paragraph({ text: text, alignment: alignment })],
    verticalAlign: docx.VerticalAlign.CENTER,
    shading,
    width: { size: width, type: docx.WidthType.DXA },
  });
}
