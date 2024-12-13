//* Указывается номер колонки по индексу и в этой колонке разделяются значения по запятой

const XLSX = require('xlsx');

function modifyColumnValues(inputFilePath, outputFilePath, columnIndex) {
    // Загрузка исходного файла
    const workbook = XLSX.readFile(inputFilePath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Получаем данные из листа
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1});

    // Проверяем, что индекс колонки в допустимом диапазоне
    if (columnIndex < 0 || columnIndex >= data[0].length) {
        throw new Error(`Column index ${columnIndex} is out of bounds`);
    }

    // Обработка значений в колонке, пропускаем первую строку (заголовки)
    data.forEach((row, index) => {
        if (index > 0 && row[columnIndex] != null) { // Пропуск первой строки и проверка на пустые ячейки
            // Триммирование начальных и конечных пробелов, замена внутренних разделителей на ', '
            let modifiedValue = row[columnIndex].toString().trim();
            modifiedValue = modifiedValue.replace(/[\s\n\r]+/g, ', ');

            // Удаление запятой, если она стоит в начале значения
            if (modifiedValue.startsWith(', ')) {
                modifiedValue = modifiedValue.substring(2);
            }

            row[columnIndex] = modifiedValue;
        }
    });

    // Создание нового листа и рабочей книги
    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, firstSheetName);

    // Сохранение нового файла
    XLSX.writeFile(newWorkbook, outputFilePath);
}

// Использование функции
modifyColumnValues('./deleteEmptyCell.xlsx', './last.xlsx', 1); // индекс с нуля
