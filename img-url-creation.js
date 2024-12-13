//* проанализировать колонку "B" взять оттуда последнее значение ( главный артикул ) и записать в колонку "F" с расширением .png

const XLSX = require('xlsx');

function extractMainArticle(inputFilePath, outputFilePath) {
    // Загрузка исходного файла
    const workbook = XLSX.readFile(inputFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Получение данных из листа
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, blankrows: false});

    // Обработка каждой строки, начиная со второй (пропускаем заголовки)
    for (let i = 1; i < data.length; i++) {
        const articles = data[i][1].split(','); // Разделение артикулов в строке
        const mainArticle = articles[articles.length - 1].trim() + '.png'; // Форматирование последнего артикула
        data[i][5] = mainArticle; // Запись в колонку "F"
    }

    // Обновление листа данными
    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    // Сохранение нового файла
    XLSX.writeFile(workbook, outputFilePath);
}

// Использование функции
extractMainArticle('./last.xlsx', './last2.xlsx');
