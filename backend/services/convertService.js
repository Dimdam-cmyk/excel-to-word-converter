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

exports.convertExcelToWord = async (filePath, discountPercentage) => {
  console.log('Начало процесса конвертации');
  console.log('Путь к файлу:', filePath);

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
        children.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: fs.readFileSync(headerImagePath),
                transformation: {
                  width: convertMillimetersToPixels(228.4),
                  height: convertMillimetersToPixels(43),
                },
              }),
            ],
            spacing: { after: 300, before: 0 },
          })
        );
        console.log('Изображение шапки добавлено');
      } catch (error) {
        console.error('Ошибка при добавлении изображения шапки:', error);
      }
    } else {
      console.log('Файл изображения шапки не найден:', headerImagePath);
    }

    // Добавляем заголовок
    children.push(
      new docx.Paragraph({
        text: "Коммерческое предложение на поставку изделий из полимербетона ARHIO",
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 0 },
        style: "Heading1"
      })
    );

    // Добавляем таблицу с данными
    const worksheet = workbook.getWorksheet(1);
    let tableRows = [];
    let totalSum = 0;

    // Добавляем заголовок таблицы
    const headerRow = new docx.TableRow({
      children: [
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Номенклатура', bold: true })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Цена, руб.', bold: true })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Площадь развёртки, м2', bold: true })] }),
      ],
    });
    tableRows.push(headerRow);

    console.log('Начало обработки строк Excel');

    // Группируем строки по значению в столбце "Наименование на фасаде"
    let groupedRows = [];
    let currentGroup = [];
    let currentName = '';

    for (let i = 2; i <= worksheet.rowCount; i++) {
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

    groupedRows.forEach(group => {
      const firstRow = group[0];
      const name = getCellValue(firstRow.getCell('A'));
      
      const mergedCell = new docx.TableCell({
        children: [new docx.Paragraph({ 
          text: name,
          alignment: docx.AlignmentType.CENTER,
          style: "italicStyle"
        })],
        rowSpan: group.length,
        verticalAlign: docx.VerticalAlign.CENTER,
      });
      
      let blockSum = 0; // Переменная для хранения суммы блока
      
      group.forEach((row, index) => {
        isEvenRow = !isEvenRow; // Меняем флаг для каждой строки
        const shading = isEvenRow ? { fill: "F2F2F2" } : undefined; // Светло-серый цвет для четных строк

        const tableRow = new docx.TableRow({
          children: index === 0 ? 
            [
              mergedCell,
              new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('B')) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('G')) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('H'))) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('I'))) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('J')), 3) })], shading }),
            ] :
            [
              new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('B')) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('G')) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('H'))) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('I'))) })], shading }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('J')), 3) })], shading }),
            ],
        });
        
        tableRows.push(tableRow);
        
        // Добавляем значение из столбца "Сумма" к общей сумме блока
        blockSum += parseFloat(getCellValue(row.getCell('I'))) || 0;
      });
      
      // Проверяем, не является ли это строкой с итоговой суммой
      const firstCellValue = getCellValue(firstRow.getCell('A'));
      if (firstCellValue.includes('Итого стоимость производства составляет')) {
        totalSum = parseFloat(getCellValue(firstRow.getCell('I'))) || 0;
        // Добавляем эту строку в таблицу без изменений
        const totalSumRow = new docx.TableRow({
          children: [
            new docx.TableCell({ children: [new docx.Paragraph({ text: firstCellValue, bold: true })], columnSpan: 5 }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalSum), bold: true, alignment: docx.AlignmentType.RIGHT })] }),
          ],
        });
        tableRows.push(totalSumRow);
        return; // Прерываем обработку этой группы
      }
      
      // Добавляем итоговую строку для блока
      const totalRow = new docx.TableRow({
        children: [
          new docx.TableCell({ 
            children: [new docx.Paragraph({ 
              text: 'Итого:',
              alignment: docx.AlignmentType.LEFT,
              bold: true
            })],
            columnSpan: 4,
          }),
          new docx.TableCell({ 
            children: [new docx.Paragraph({ 
              text: formatNumber(blockSum),
              alignment: docx.AlignmentType.CENTER,
              bold: true
            })],
          }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: '' })] }),
        ],
        shading: { fill: "D9D9D9" }, // Серый цвет для строки "Итого"
      });
      
      tableRows.push(totalRow);
      isEvenRow = false; // Сбрасываем флаг после итоговой строки
    });

    console.log(`Обработано ${tableRows.length} строк`);

    // Удаляем последнюю строку из tableRows, так как она содержит итоговую сумму
    tableRows.pop();

    // Добавляем таблицу
    const table = new docx.Table({
      rows: tableRows,
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      },
    });

    children.push(table);

    // Добавляем итоговую сумму
    children.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Итого стоимость производства составляет ${formatNumber(totalSum)} руб.`,
            bold: true,
          }),
        ],
        spacing: { before: 400, after: 400 },
      })
    );

    // Добавляем информацию о скидке, если она применяется
    let discountedTotal = totalSum;
    if (discountPercentage) {
      const discountAmount = totalSum * (discountPercentage / 100);
      discountedTotal = totalSum - discountAmount;

      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `Цена со скидкой ${discountPercentage}%: ${formatNumber(discountedTotal)} руб.`,
              bold: true,
            }),
          ],
          spacing: { before: 200, after: 200 },
        })
      );

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
          children.push(
            new docx.Paragraph({
              children: [
                new docx.ImageRun({
                  data: fs.readFileSync(imagePath),
                  transformation: {
                    width: convertMillimetersToPixels(277),
                    height: convertMillimetersToPixels(190),
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

    // Добвляем одну секцию со всем содержимым и колонтитулами
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
              text: "Предложение действительно 25 дней. Предложение является предварительным. Для окончательного расчета требуется проектирование.",
              alignment: docx.AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children: children,
    });

    console.log('Тблица, изображения и колонтитулы добавлены в документ Word');

    console.log('Начало создания буфера документа Word');
    const buffer = await docx.Packer.toBuffer(doc);
    console.log('Буфер документа Word создан');

    fs.unlinkSync(filePath);
    console.log('Временный файл Excel удален');

    return buffer;
  } catch (error) {
    console.error('Ошибка в процессе конвертации:', error);
    console.error('Стек вызовов:', error.stack);
    
    // Добавьте больше информации об ошибке
    if (error instanceof ExcelJS.Error) {
      console.error('Ошибка ExcelJS:', error.message);
    } else if (error instanceof docx.Error) {
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
