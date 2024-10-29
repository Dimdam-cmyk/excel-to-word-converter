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

exports.convertExcelToWord = async (filePath, discountPercentage, makeShortVersion, originalFileName) => {
  console.log('Начало процесса конвертации');
  console.log('Путь к файлу:', filePath);
  console.log('Оригинальное имя файла:', originalFileName);

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
    // Получаем имя файла без расширения и первые 10 символов
    const fileId = originalFileName.replace('.xlsx', '').slice(0, 10);

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
      children: [
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true })], alignment: docx.AlignmentType.CENTER }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Номенклатура', bold: true })], alignment: docx.AlignmentType.CENTER }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true })], alignment: docx.AlignmentType.CENTER }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Цена, руб.', bold: true })], alignment: docx.AlignmentType.CENTER }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Сумма, руб.', bold: true })], alignment: docx.AlignmentType.CENTER }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Площадь развёртки, м2', bold: true })], alignment: docx.AlignmentType.CENTER }),
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
        const headerRow = new docx.TableRow({
          children: [
            new docx.TableCell({ children: [new docx.Paragraph({ text: 'Наименование на фасаде', bold: true })], alignment: docx.AlignmentType.CENTER }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: 'Сумма', bold: true })], alignment: docx.AlignmentType.CENTER }),
          ],
        });
        tableRows.push(headerRow);

        groupedRows.forEach(group => {
          const firstRow = group[0];
          const name = getCellValue(firstRow.getCell('A'));
          
          // Пропускаем строку "Итого стоимость работ составляет"
          if (name.includes('Итого стоимость производства составляет')) {
            return;
          }
          
          let blockSum = 0;
          
          group.forEach(row => {
            blockSum += parseFloat(getCellValue(row.getCell('I'))) || 0;
          });

          const tableRow = new docx.TableRow({
            children: [
              new docx.TableCell({ children: [new docx.Paragraph({ text: name, alignment: docx.AlignmentType.LEFT })], verticalAlign: docx.VerticalAlign.CENTER }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(blockSum), alignment: docx.AlignmentType.RIGHT })], verticalAlign: docx.VerticalAlign.CENTER }),
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
                  }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('B')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('G')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('H'))), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('I'))), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('J')), 3), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                ] :
                [
                  new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('B')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('G')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('H'))), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('I'))), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                  new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(getCellValue(row.getCell('J')), 3), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
                ],
            });
            
            tableRows.push(tableRow);
            
            blockSum += parseFloat(getCellValue(row.getCell('I'))) || 0;
          });
          
          // Добавляем итоговую строку для блока
          const blockTotalRow = new docx.TableRow({
            children: [
              new docx.TableCell({ children: [new docx.Paragraph({ text: 'Итого:', bold: true })], columnSpan: 4, alignment: docx.AlignmentType.RIGHT }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(blockSum), bold: true, alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
              new docx.TableCell({ children: [new docx.Paragraph({ text: '' })], verticalAlign: docx.VerticalAlign.CENTER }),
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

    // Добавляем итоговую строку
    const totalSumRow = new docx.TableRow({
      children: makeShortVersion ? [
        new docx.TableCell({ 
          children: [new docx.Paragraph({ 
            text: 'Итого стоимость производства составляет', 
            bold: true,
            alignment: docx.AlignmentType.LEFT 
          })],
        }),
        new docx.TableCell({ 
          children: [new docx.Paragraph({ 
            text: formatNumber(totalSum), 
            bold: true, 
            alignment: docx.AlignmentType.CENTER 
          })],
          verticalAlign: docx.VerticalAlign.CENTER 
        }),
      ] : [
        new docx.TableCell({ children: [new docx.Paragraph({ text: 'Итого:', bold: true })], columnSpan: 4, alignment: docx.AlignmentType.LEFT }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalSum), bold: true, alignment: docx.AlignmentType.RIGHT })], verticalAlign: docx.VerticalAlign.CENTER }),
        new docx.TableCell({ children: [new docx.Paragraph({ text: '' })], verticalAlign: docx.VerticalAlign.CENTER }),
      ],
    });

    // Применяем заливку к ячейкам итоговой строки
    if (totalSumRow.children) {
      totalSumRow.children.forEach(cell => {
        if (cell) {
          cell.shading = { fill: "FFE699" };
        }
      });
    }

    tableRows.push(totalSumRow);

    // Добавляем таблицу
    const table = new docx.Table({
      rows: tableRows,
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      },
    });

    children.push(table);

    // Добавляем новую таблицу "Стоимость форм и заливки"
    children.push(
      new docx.Paragraph({
        text: "Стоимость форм и заливки",
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300, before: 300 },
        style: "Heading1"
      })
    );

    const izdeliyaWorksheet = workbook.getWorksheet('Изделия');
    if (izdeliyaWorksheet) {
      console.log('Найден лист "Изделия"');
      let tableRows = [];
      let isEvenRow = false;
      let totalH = 0, totalI = 0, totalK = 0, totalL = 0, totalN = 0;

      // Добавляем заголовок таблицы
      const headerRow = new docx.TableRow({
        children: [
          new docx.TableCell({ children: [new docx.Paragraph({ text: '№ п/п', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Номенклатура', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Площадь развертки изделия, м2', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Кол-во изделий, шт.', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Площадь развертки общая, м2', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Масса общая, кг', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Кол-во форм, шт.', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Стоимость формы за м², руб.', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Стоимость форм для изделий, руб.', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Стоимость заливки за м², руб.', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: 'Стоимость за единицу, руб.', bold: true })], alignment: docx.AlignmentType.CENTER }),
        ],
        tableHeader: true,
      });
      tableRows.push(headerRow);

      console.log('Начало обработки строк листа "Изделия"');
      for (let i = 2; i <= izdeliyaWorksheet.rowCount; i++) {
        const row = izdeliyaWorksheet.getRow(i);
        const cellA = getCellValue(row.getCell('A'));
        if (!cellA) {
          console.log(`Достигнут конец данных на строке ${i}`);
          break;
        }

        console.log(`Обработка строки ${i}`);
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

        const tableRow = new docx.TableRow({
          children: [
            new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('A')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
            new docx.TableCell({ children: [new docx.Paragraph({ text: getCellValue(row.getCell('B')), alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER, shading }),
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
      }

      console.log('Добавление итоговой строки');
      // Добавляем итоговую строку
      const totalTableRow = new docx.TableRow({
        children: [
          new docx.TableCell({ children: [new docx.Paragraph({ text: '', bold: true })], columnSpan: 3 }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalH, 0), bold: true, alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalI, 2), bold: true, alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalK, 2), bold: true, alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalL, 0), bold: true, alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: '', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: formatNumber(totalN * (1 - discountPercentage / 100), 2), bold: true, alignment: docx.AlignmentType.CENTER })], verticalAlign: docx.VerticalAlign.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: '', bold: true })], alignment: docx.AlignmentType.CENTER }),
          new docx.TableCell({ children: [new docx.Paragraph({ text: '', bold: true })], alignment: docx.AlignmentType.CENTER }),
        ],
        height: {
          value: 400,
          rule: docx.HeightRule.ATLEAST,
        },
      });

      // Применяем жирный шрифт для всех ячеек в итоговой строке
      if (totalTableRow.children) {
        totalTableRow.children.forEach(cell => {
          if (cell && cell.children) {
            cell.children.forEach(paragraph => {
              if (paragraph && paragraph.options) {
                paragraph.options.bold = true;
              }
            });
          }
        });
      }

      tableRows.push(totalTableRow);

      console.log('Создание таблицы');
      const table = new docx.Table({
        rows: tableRows,
        width: {
          size: 100,
          type: docx.WidthType.PERCENTAGE,
        },
      });

      children.push(table);
      console.log('Таблица "Стоимость форм и заливки" успешно добавлена');
    } else {
      console.log('Лист "Изделия" не найден');
    }

    console.log('Таблица "Стоимость форм и заливки" добавлена в документ Word');

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
              text: "Предложение действительно 25 дней. Расчет является предварительным. Для окончательног расчета требуется проектирование.",
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
