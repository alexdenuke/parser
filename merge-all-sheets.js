//* Объединяет все листы в эксель файле в один файл

const XLSX = require('xlsx');

function combineSheets(inputFilePath, outputFilePath) {
    // Загрузка исходного файла
    const workbook = XLSX.readFile(inputFilePath);
    let combinedData = [];

    // Объединение данных всех листов
    workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        // Получение данных из листа и добавление в общий массив
        const data = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        if (combinedData.length === 0) {
            // Добавление заголовков только один раз
            combinedData = data;
        } else {
            // Добавление данных, исключая повтор заголовков
            combinedData.push(...data.slice(1));
        }
    });

    // Создание нового рабочего книги и листа с объединенными данными
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(combinedData);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Combined');

    // Сохранение нового файла
    XLSX.writeFile(newWorkbook, outputFilePath);
}

// входящий и выходяший файл
combineSheets('./partsparts_category.xlsx', './oneSheet.xlsx');
