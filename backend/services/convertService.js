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

exports.convertExcelToWord = async (filePath, discountPercentage, makeShortVersion, originalFileName, includeVAT = false, rbtChecked = false, extraExpenses = [], rbtDiscount = null, pricePerSqm = 14000, subsystemPercentage = 15) => {
  console.log('=== Начало конвертации ===');
  console.log('Параметры:');
  console.log('- filePath:', filePath);
  console.log('- discountPercentage:', discountPercentage);
  console.log('- makeShortVersion:', makeShortVersion);
  console.log('- originalFileName:', originalFileName);
  console.log('- includeVAT:', includeVAT);
  console.log('- rbtChecked:', rbtChecked);
  console.log('- extraExpenses:', extraExpenses);
  console.log('- rbtDiscount:', rbtDiscount);
  console.log('- pricePerSqm:', pricePerSqm);
  console.log('- subsystemPercentage:', subsystemPercentage);

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
        children: [
          new docx.TextRun({
            text: `Коммерческое предложение на поставку изделий из полимербетона ARHIO по проекту ${fileId}`,
            bold: true,
            size: 28, // 14pt
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 0 },
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
          children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER, 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" } 
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
      ] : [
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER, 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" } // Добавляем серую заливку для заголовка
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Номенклатура', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Цена, руб.', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true, alignment: docx.AlignmentType.CENTER })], 
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ text: 'Площадь развёртки, м2', bold: true, alignment: docx.AlignmentType.CENTER })], 
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
                alignment: docx.AlignmentType.CENTER
              })
            ],
            alignment: docx.AlignmentType.CENTER
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
          
          // Добавляем итоговую строку для блока с желтой заливкой
          const blockTotalRow = new docx.TableRow({
            children: [
              new docx.TableCell({ 
                children: [itogoParagraph], 
                columnSpan: 3,
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "DDE8F6" } // Светло-голубая заливка
              }),
              new docx.TableCell({ 
                children: [emptyBoldParagraph], 
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "DDE8F6" } // Светло-голубая заливка
              }),
              new docx.TableCell({ 
                children: [totalSumParagraph], 
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "DDE8F6" } // Светло-голубая заливка
              }),
              new docx.TableCell({ 
                children: [totalAreaParagraph], 
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "DDE8F6" } // Светло-голубая заливка
              }),
            ],
          });
          tableRows.push(blockTotalRow);
          
          // Добавляем пустую строку после итогового блока
          const emptyRow = new docx.TableRow({
            children: makeShortVersion ? [
              new docx.TableCell({ children: [new docx.Paragraph({ text: '' })], columnSpan: 2 }),
            ] : [
              new docx.TableCell({ children: [new docx.Paragraph({ text: '' })], columnSpan: 6 }),
            ]
          });
          tableRows.push(emptyRow);
          
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
          shading: { fill: "DDE8F6" } // Светло-голубая заливка
        }),
        new docx.TableCell({ 
          children: [totalSumParagraph],
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" } // Светло-голубая заливка
        }),
      ] : [
        new docx.TableCell({ 
          children: [totalProjectParagraph], 
          columnSpan: 3, 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" } // Светло-голубая заливка
        }),
        new docx.TableCell({ 
          children: [emptyBoldTotal], 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" } // Светло-голубая заливка
        }),
        new docx.TableCell({ 
          children: [totalSumParagraph], 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" } // Светло-голубая заливка
        }),
        new docx.TableCell({ 
          children: [totalAreaParagraph], 
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" } // Светло-голубая заливка
        }),
      ],
    });

    // Текст "ИТОГО по проекту:" выравниваем по центру
    totalProjectParagraph.alignment = docx.AlignmentType.CENTER;

    // Применяем заливку к ячейкам итоговой строки
    if (totalSumRow.children) {
      totalSumRow.children.forEach(cell => {
        if (cell) {
          cell.shading = { fill: "DDE8F6" }; // Светло-голубая заливка для итоговой строки
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
        tableRows[0].children[0].width = { size: 2500, type: docx.WidthType.DXA };
        
        // Номенклатура - средний, увеличиваем ширину
        tableRows[0].children[1].width = { size: 3000, type: docx.WidthType.DXA };
        
        // Кол-во изделий, шт.
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
        children: [
          new docx.TextRun({
            text: "Стоимость форм и заливки",
            bold: true,
            size: 28, // 14pt
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 300 },
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
            children: [new docx.Paragraph({ text: '№ п/п', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Номенклатура', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Площадь развертки изделия, м2', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Площадь развертки общая, м2', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Масса общая, кг', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Кол-во форм, шт.', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость формы за м², руб.', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость форм для изделий, руб.', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость заливки за м², руб.', bold: true, alignment: docx.AlignmentType.CENTER })],
            alignment: docx.AlignmentType.CENTER,
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "D3D3D3" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Стоимость за единицу, руб.', bold: true, alignment: docx.AlignmentType.CENTER })],
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
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedM, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading: { fill: "DDE8F6" } }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedN, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(discountedO, 2), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading: { fill: "DDE8F6" } }),
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
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalH, 0),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalI, 2),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalK, 2),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalL, 0),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: '',
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: formatNumber(totalN * (1 - discountPercentage / 100), 2),
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: '',
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({
              text: '',
              bold: true,
              alignment: docx.AlignmentType.CENTER
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "DDE8F6" }
          }),
        ],
        height: {
          value: 400,
          rule: docx.HeightRule.ATLEAST,
        },
      });

      // Добавляем итоговую строку
      tableRows.push(totalTableRow);

      // Добавляем пустую строку после итоговой строки
      const emptySecondTableRow = new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ text: '' })],
            columnSpan: 11
          })
        ]
      });
      tableRows.push(emptySecondTableRow);

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
        
        // Номенклатура - самый широкий (увеличиваю ширину)
        tableRows[0].children[1].width = { size: 3000, type: docx.WidthType.DXA };
        
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
    // Константы НДС: используем ставку 22% и формулу выделения НДС из суммы "с НДС"
    // VAT = Gross * rate / (1 + rate)
    const VAT_RATE = 0.22;
    const VAT_PERCENT = 22;
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Итого стоимость производства составляет ${formatNumber(totalSum)} руб. (${numberToWordsRubles(Math.round(totalSum))}${includeVAT ? ` включая НДС ${VAT_PERCENT}%` : ' без НДС'})`,
            bold: true,
          }),
        ],
        spacing: { before: 400, after: 200 },
      })
    );

    // Добавляем сумму НДС если чекбокс активирован
    if (includeVAT) {
      // Рассчитываем НДС (22% от общей суммы)
      const vatAmount = totalSum * VAT_RATE / (1 + VAT_RATE);
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `В том числе НДС ${VAT_PERCENT}%: ${formatNumber(vatAmount)} руб.`,
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
              text: `Цена со скидкой ${discountPercentage}%: ${formatNumber(discountedTotal)} руб. (${numberToWordsRubles(Math.round(discountedTotal))}${includeVAT ? ` включая НДС ${VAT_PERCENT}%` : ' без НДС'})`,
              bold: true,
            }),
          ],
          spacing: { before: 200, after: 200 },
        })
      );

      // Добавляем сумму НДС для цены со скидкой, если чекбокс активирован
      if (includeVAT) {
        // Рассчитываем НДС (22% от суммы со скидкой)
        const vatAmountWithDiscount = discountedTotal * VAT_RATE / (1 + VAT_RATE);
        children.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `В том числе НДС ${VAT_PERCENT}%: ${formatNumber(vatAmountWithDiscount)} руб.`,
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

    // Добавляем изображение подписи с контактами
    const signaturePath = path.join(__dirname, '../assets/signature.png');
    if (fs.existsSync(signaturePath)) {
      try {
        const signatureBuffer = fs.readFileSync(signaturePath);
        // Конвертируем сантиметры в миллиметры (1 см = 10 мм)
        const signatureWidthMm = 121.3; // 12.13 см
        const signatureHeightMm = 41.7; // 4.17 см
        
        children.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: signatureBuffer,
                transformation: {
                  width: convertMillimetersToPixels(signatureWidthMm),
                  height: convertMillimetersToPixels(signatureHeightMm),
                },
              }),
            ],
            alignment: docx.AlignmentType.LEFT,
            spacing: { before: 300, after: 300 }, // Добавляем отступы для лучшего вида
          })
        );
        console.log('Изображение подписи с контактами добавлено');
      } catch (error) {
        console.error('Ошибка при добавлении изображения подписи:', error);
      }
    } else {
      console.log('Файл изображения подписи не найден:', signaturePath);
    }

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

    // Если включен РБТ, создаем дополнительный Word документ для монтажа
    let mountingWordBuffer = null;
    if (rbtChecked) {
      console.log('Создание Word документа для монтажного коммерческого предложения');
      
      // Рассчитываем общую площадь для монтажа
      let totalArea = 0;
      let rowCount = 0;
      if (groupedRows && Array.isArray(groupedRows)) {
        groupedRows.forEach((group, groupIndex) => {
          group.forEach((row, rowIndex) => {
            const nomenclature = getCellValue(row.getCell('B'));
            const area = parseFloat(getCellValue(row.getCell('J'))) || 0;
            
            // Пропускаем строки с итоговыми значениями или пустыми номенклатурами
            if (!nomenclature || nomenclature.trim() === '' || 
                nomenclature.includes('Итого') || 
                nomenclature.includes('итого') ||
                nomenclature.includes('ИТОГО')) {
              console.log(`Пропускаем итоговую строку: ${nomenclature}, площадь: ${area}`);
              return;
            }
            
            console.log(`Группа ${groupIndex}, строка ${rowIndex}: ${nomenclature}, площадь: ${area}`);
            totalArea += area;
            rowCount++;
          });
        });
      }
      console.log(`Всего обработано строк: ${rowCount}, общая площадь: ${totalArea}`);
      const totalMountingCost = totalArea * pricePerSqm;
      console.log(`Цена за м.кв.: ${pricePerSqm}, общая стоимость монтажа: ${totalMountingCost}`);
      
      mountingWordBuffer = await createMountingWord(workbook, groupedRows, extraExpenses, rbtDiscount, originalFileName, pricePerSqm, totalMountingCost, subsystemPercentage);
      console.log('Word документ для монтажа создан');
    }

    fs.unlinkSync(filePath);
    console.log('Временный файл Excel удален');
    
    // Возвращаем оба файла если есть РБТ
    if (mountingWordBuffer) {
      return { wordBuffer: buffer, pdfBuffer: mountingWordBuffer };
    }
    
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

// Функция для создания Word документа с монтажным коммерческим предложением
async function createMountingWord(workbook, groupedRows, extraExpenses, rbtDiscount, originalFileName, pricePerSqm = 14000, totalMountingCost = 0, subsystemPercentage = 15) {
  try {
    const doc = new docx.Document({
      styles: {
        paragraphStyles: [
          {
            id: "mountingTitleStyle",
            name: "Mounting Title Style",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            run: {
              size: 28, // 14 пунктов
              bold: true,
            },
            paragraph: {
              alignment: docx.AlignmentType.CENTER,
            },
          },
          {
            id: "mountingHeaderStyle",
            name: "Mounting Header Style",
            basedOn: "Normal",
            run: {
              size: 22, // 11 пунктов
              bold: true,
            },
          },
        ],
      },
      sections: []
    });

    const children = [];
    
    // Получаем имя файла без расширения
    const fileId = originalFileName ? path.parse(originalFileName).name : 'unknown';

    // Добавляем заголовок
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Коммерческое предложение на монтаж фасадного декора`,
            bold: true,
            size: 28,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 0 },
      })
    );

    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Проект: ${fileId}`,
            bold: true,
            size: 24,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 400, before: 0 },
      })
    );

    // Создаем таблицу для монтажа
    let tableRows = [];

    // Заголовок таблицы
    const headerRow = new docx.TableRow({
      children: [
        new docx.TableCell({
          children: [new docx.Paragraph({ text: 'Номенклатура', bold: true, alignment: docx.AlignmentType.CENTER })],
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({
          children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true, alignment: docx.AlignmentType.CENTER })],
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({
          children: [new docx.Paragraph({ text: 'Площадь развертки, м²', bold: true, alignment: docx.AlignmentType.CENTER })],
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
        new docx.TableCell({
          children: [new docx.Paragraph({ text: 'Стоимость монтажа, руб.', bold: true, alignment: docx.AlignmentType.CENTER })],
          alignment: docx.AlignmentType.CENTER,
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "D3D3D3" }
        }),
      ],
    });
    tableRows.push(headerRow);

    // Добавляем данные
    let isEvenRow = false;
    if (groupedRows && Array.isArray(groupedRows)) {
      groupedRows.forEach(group => {
        group.forEach(row => {
          const nomenclature = getCellValue(row.getCell('B'));
          
          // Пропускаем строки с итоговыми значениями или пустыми номенклатурами
          if (!nomenclature || nomenclature.trim() === '' || 
              nomenclature.includes('Итого') || 
              nomenclature.includes('итого') ||
              nomenclature.includes('ИТОГО')) {
            return;
          }
          
          isEvenRow = !isEvenRow;
          const shading = isEvenRow ? { fill: "F2F2F2" } : undefined;
          
          const quantity = getCellValue(row.getCell('G'));
          const area = parseFloat(getCellValue(row.getCell('J'))) || 0;
          const mountingCost = area * pricePerSqm; // площадь × цена за м.кв.
          
          const tableRow = new docx.TableRow({
            children: [
              new docx.TableCell({
                children: [new docx.Paragraph({ text: nomenclature, alignment: docx.AlignmentType.LEFT })],
                verticalAlign: docx.VerticalAlign.CENTER,
                shading
              }),
              new docx.TableCell({
                children: [new docx.Paragraph({ text: quantity, alignment: docx.AlignmentType.CENTER })],
                verticalAlign: docx.VerticalAlign.CENTER,
                shading
              }),
              new docx.TableCell({
                children: [new docx.Paragraph({ text: formatNumber(area, 2), alignment: docx.AlignmentType.CENTER })],
                verticalAlign: docx.VerticalAlign.CENTER,
                shading
              }),
              new docx.TableCell({
                children: [new docx.Paragraph({ text: formatNumber(mountingCost, 2), alignment: docx.AlignmentType.CENTER })],
                verticalAlign: docx.VerticalAlign.CENTER,
                shading
              }),
            ],
          });
          tableRows.push(tableRow);
        });
      });
    }

    // Создаем таблицу
    const mountingTable = new docx.Table({
      rows: tableRows,
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      },
    });

    children.push(mountingTable);

    // Добавляем небольшой отступ
    children.push(
      new docx.Paragraph({
        children: [],
        spacing: { before: 200, after: 100 },
      })
    );

    // Добавляем строку "Итого по монтажу изделий" в виде таблицы для единого стиля
    const totalMountingRow = new docx.TableRow({
      children: [
        new docx.TableCell({
          children: [new docx.Paragraph({ 
            children: [new docx.TextRun({ text: 'Итого по монтажу изделий:', bold: true })],
            alignment: docx.AlignmentType.RIGHT 
          })],
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "E8F4FD" },
          width: { size: 70, type: docx.WidthType.PERCENTAGE }
        }),
        new docx.TableCell({
          children: [new docx.Paragraph({ 
            children: [new docx.TextRun({ text: formatNumber(totalMountingCost, 2) + ' руб.', bold: true })],
            alignment: docx.AlignmentType.CENTER 
          })],
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "E8F4FD" },
          width: { size: 30, type: docx.WidthType.PERCENTAGE }
        }),
      ],
    });

    const totalMountingTable = new docx.Table({
      rows: [totalMountingRow],
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      },
    });

    children.push(totalMountingTable);

    // Добавляем дополнительные расходы как таблицу если есть
    let extraExpensesTotal = 0;
    
    // Автоматически добавляем алюминиевую подсистему (задаваемый процент от стоимости монтажа)
    const aluminumCost = totalMountingCost * (subsystemPercentage / 100);
    extraExpensesTotal += aluminumCost;
    
    // Проверяем, есть ли дополнительные расходы от пользователя
    const hasUserExpenses = extraExpenses && extraExpenses.length > 0 && extraExpenses.some(expense => {
      if (typeof expense === 'string') return expense && expense.trim();
      if (expense && typeof expense === 'object') return expense.name && expense.name.trim();
      return false;
    });
    
    // Всегда показываем таблицу расходов, так как есть автоматическая строка с алюминием
    const hasExpenses = true;

    if (hasExpenses) {
        children.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: "Дополнительные расходы",
                bold: true,
                size: 24,
              }),
            ],
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 400, after: 300 },
          })
        );

        // Создаем таблицу для дополнительных расходов
        let expenseTableRows = [];
        
        // Заголовок таблицы
        const expenseHeaderRow = new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [new docx.Paragraph({ text: 'Наименование расхода', bold: true, alignment: docx.AlignmentType.CENTER })],
              alignment: docx.AlignmentType.CENTER,
              verticalAlign: docx.VerticalAlign.CENTER,
              shading: { fill: "D3D3D3" },
              width: { size: 70, type: docx.WidthType.PERCENTAGE }
            }),
            new docx.TableCell({
              children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true, alignment: docx.AlignmentType.CENTER })],
              alignment: docx.AlignmentType.CENTER,
              verticalAlign: docx.VerticalAlign.CENTER,
              shading: { fill: "D3D3D3" },
              width: { size: 30, type: docx.WidthType.PERCENTAGE }
            }),
          ],
        });
        expenseTableRows.push(expenseHeaderRow);

        // Добавляем автоматическую строку с алюминиевой подсистемой
        let expenseRowIndex = 0;
        
        // Первая строка - алюминиевая подсистема (автоматически)
        expenseRowIndex++;
        const aluminumRow = new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [new docx.Paragraph({ text: 'Алюминиевая подсистема и крепежи', alignment: docx.AlignmentType.LEFT })],
              verticalAlign: docx.VerticalAlign.CENTER,
              shading: { fill: "E8F4FD" } // Светло-голубой фон для автоматической строки
            }),
            new docx.TableCell({
              children: [new docx.Paragraph({ 
                text: formatNumber(aluminumCost, 2), 
                alignment: docx.AlignmentType.CENTER 
              })],
              verticalAlign: docx.VerticalAlign.CENTER,
              shading: { fill: "E8F4FD" }
            }),
          ],
        });
        expenseTableRows.push(aluminumRow);
        
        // Добавляем пользовательские расходы
        if (hasUserExpenses) {
          extraExpenses.forEach((expense, index) => {
          // Поддержка старого формата (строки) и нового формата (объекты)
          let name, amount;
          if (typeof expense === 'string') {
            name = expense;
            amount = 0;
          } else if (expense && typeof expense === 'object') {
            name = expense.name || `Доп. расход ${index + 1}`;
            amount = parseFloat(expense.amount) || 0;
          } else {
            return;
          }
          
          if (name && name.trim()) {
            extraExpensesTotal += amount;
            expenseRowIndex++;
            
            const isEvenExpenseRow = expenseRowIndex % 2 === 0;
            const expenseShading = isEvenExpenseRow ? { fill: "F2F2F2" } : undefined;
            
            const expenseRow = new docx.TableRow({
              children: [
                new docx.TableCell({
                  children: [new docx.Paragraph({ text: name, alignment: docx.AlignmentType.LEFT })],
                  verticalAlign: docx.VerticalAlign.CENTER,
                  shading: expenseShading
                }),
                new docx.TableCell({
                  children: [new docx.Paragraph({ 
                    text: amount > 0 ? formatNumber(amount, 2) : '—', 
                    alignment: docx.AlignmentType.CENTER 
                  })],
                  verticalAlign: docx.VerticalAlign.CENTER,
                  shading: expenseShading
                }),
              ],
            });
            expenseTableRows.push(expenseRow);
          }
        });
        }
        
        // Итоговая строка для расходов
        if (extraExpensesTotal > 0) {
          const expenseTotalRow = new docx.TableRow({
            children: [
              new docx.TableCell({
                children: [new docx.Paragraph({ 
                  children: [new docx.TextRun({ text: 'Итого дополнительные расходы:', bold: true })],
                  alignment: docx.AlignmentType.RIGHT 
                })],
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "DDE8F6" }
              }),
              new docx.TableCell({
                children: [new docx.Paragraph({ 
                  children: [new docx.TextRun({ text: formatNumber(extraExpensesTotal, 2), bold: true })],
                  alignment: docx.AlignmentType.CENTER 
                })],
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "DDE8F6" }
              }),
            ],
          });
          expenseTableRows.push(expenseTotalRow);
        }

        // Создаем таблицу расходов
        const expenseTable = new docx.Table({
          rows: expenseTableRows,
          width: {
            size: 100,
            type: docx.WidthType.PERCENTAGE,
          },
        });

        children.push(expenseTable);
    }

    // Итоговые расчеты в виде красивой таблицы
    const totalWithExpenses = totalMountingCost + extraExpensesTotal;
    let finalTotal = totalWithExpenses;
    
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "Итоговый расчет",
            bold: true,
            size: 24,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 400, after: 300 },
      })
    );

    // Создаем таблицу итогового расчета
    let totalTableRows = [];
    
    // Стоимость монтажа
    const mountingRow = new docx.TableRow({
      children: [
        new docx.TableCell({
          children: [new docx.Paragraph({ text: 'Стоимость монтажа', alignment: docx.AlignmentType.LEFT })],
          verticalAlign: docx.VerticalAlign.CENTER,
          width: { size: 70, type: docx.WidthType.PERCENTAGE }
        }),
        new docx.TableCell({
          children: [new docx.Paragraph({ text: formatNumber(totalMountingCost, 2) + ' руб.', alignment: docx.AlignmentType.CENTER })],
          verticalAlign: docx.VerticalAlign.CENTER,
          width: { size: 30, type: docx.WidthType.PERCENTAGE }
        }),
      ],
    });
    totalTableRows.push(mountingRow);

    // Дополнительные расходы если есть
    if (extraExpensesTotal > 0) {
      const expensesRow = new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ text: 'Итого дополнительные расходы', alignment: docx.AlignmentType.LEFT })],
            verticalAlign: docx.VerticalAlign.CENTER
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: formatNumber(extraExpensesTotal, 2) + ' руб.', alignment: docx.AlignmentType.CENTER })],
            verticalAlign: docx.VerticalAlign.CENTER
          }),
        ],
      });
      totalTableRows.push(expensesRow);
      
      const subtotalRow = new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ 
              children: [new docx.TextRun({ text: 'Итого монтаж под ключ', bold: true })],
              alignment: docx.AlignmentType.LEFT 
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "F2F2F2" }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ 
              children: [new docx.TextRun({ text: formatNumber(totalWithExpenses, 2) + ' руб.', bold: true })],
              alignment: docx.AlignmentType.CENTER 
            })],
            verticalAlign: docx.VerticalAlign.CENTER,
            shading: { fill: "F2F2F2" }
          }),
        ],
      });
      totalTableRows.push(subtotalRow);
    }
    
    // Скидка если есть
    if (rbtDiscount && rbtDiscount > 0) {
      const discountAmount = totalWithExpenses * (rbtDiscount / 100);
      finalTotal = totalWithExpenses - discountAmount;
      
      const discountRow = new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ text: `Скидка ${rbtDiscount}%`, alignment: docx.AlignmentType.LEFT })],
            verticalAlign: docx.VerticalAlign.CENTER
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ text: '- ' + formatNumber(discountAmount, 2) + ' руб.', alignment: docx.AlignmentType.CENTER })],
            verticalAlign: docx.VerticalAlign.CENTER
          }),
        ],
      });
      totalTableRows.push(discountRow);
    }
    
    // Финальная итоговая строка
    const finalRow = new docx.TableRow({
      children: [
        new docx.TableCell({
          children: [new docx.Paragraph({ 
            children: [new docx.TextRun({ text: 'Итого монтаж под ключ', bold: true, size: 24 })],
            alignment: docx.AlignmentType.LEFT 
          })],
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" }
        }),
        new docx.TableCell({
          children: [new docx.Paragraph({ 
            children: [new docx.TextRun({ text: formatNumber(finalTotal, 2) + ' руб.', bold: true, size: 24 })],
            alignment: docx.AlignmentType.CENTER 
          })],
          verticalAlign: docx.VerticalAlign.CENTER,
          shading: { fill: "DDE8F6" }
        }),
      ],
    });
    totalTableRows.push(finalRow);

    // Создаем итоговую таблицу
    const totalTable = new docx.Table({
      rows: totalTableRows,
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      },
    });

    children.push(totalTable);
    
    // Добавляем сумму прописью
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `(${numberToWordsRubles(Math.round(finalTotal))})`,
            italics: true,
            size: 22,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 200, after: 400 },
      })
    );

    // Добавляем примечания

    // Добавляем таблицу с примечаниями
    const notesTableRows = [];
    
    // Заголовок примечаний
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "В стоимость включено:",
            bold: true,
            size: 22,
          }),
        ],
        spacing: { before: 400, after: 200 },
      })
    );

    const includedText = `Уборка мусора в контейнер.
Доставка всех материалов для монтажа на объект (подсистема, расходные материалы) за исключением изделий.
Доставка, аренда и сборка лесов.`;

    children.push(
      new docx.Paragraph({
        text: includedText,
        spacing: { before: 100, after: 300 },
      })
    );

    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "В стоимость НЕ включено:",
            bold: true,
            size: 22,
          }),
        ],
        spacing: { before: 200, after: 200 },
      })
    );

    const notIncludedText = `Работы по возведению несущих конструкций из кирпича и горячекатанного металла.
Сварочные работы.
Малярные работы.
Электромонтажные работы.
Земельные работы.`;

    children.push(
      new docx.Paragraph({
        text: notIncludedText,
        spacing: { before: 100, after: 300 },
      })
    );

    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "Примечание:",
            bold: true,
            size: 22,
          }),
        ],
        spacing: { before: 200, after: 200 },
      })
    );

    const noteText = `В стоимость монтажа входят все работы связанные с установкой изделий на фасад, дополнительными работами считаются работы которые возникают в процессе монтажа по причине вмешательства смежных бригад или просьб заказчика.`;

    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: noteText,
            italics: true,
          }),
        ],
        spacing: { before: 100, after: 300 },
      })
    );

    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "Форма оплаты - наличный расчет",
            bold: true,
            underline: {},
            size: 22,
          }),
        ],
        spacing: { before: 400, after: 200 },
      })
    );

    // Получаем текущую дату
    const currentDate = new Date().toLocaleDateString('ru-RU');

    // Добавляем секцию с вертикальной ориентацией
    doc.addSection({
      properties: {
        page: {
          size: {
            width: convertMillimetersToTwip(210), // A4 ширина в портретной ориентации
            height: convertMillimetersToTwip(297), // A4 высота в портретной ориентации
          },
          orientation: docx.PageOrientation.PORTRAIT, // Вертикальная ориентация
          margins: {
            top: convertMillimetersToTwip(20),
            right: convertMillimetersToTwip(15),
            bottom: convertMillimetersToTwip(20),
            left: convertMillimetersToTwip(15),
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
              text: "Предложение действительно 25 дней. Расчет является предварительным.",
              alignment: docx.AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children: children,
    });

    console.log('Создание буфера документа Word для монтажа');
    const buffer = await docx.Packer.toBuffer(doc);
    console.log('Буфер документа Word для монтажа создан');
    
    return buffer;
  } catch (error) {
    console.error('Ошибка при создании Word документа для монтажа:', error);
    throw error;
  }
}

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
