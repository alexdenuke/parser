const XLSX = require('xlsx');

function cleanAndSaveExcel(inputFilePath, outputFilePath) {
    // Загрузка исходного файла
    const workbook = XLSX.readFile(inputFilePath);
    
    // Предполагаем, что работаем только с первым листом
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Конвертируем лист в массив объектов
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1});
    
    // Фильтрация данных, удаляем строки, где нет данных в колонке "B" (индекс 1)
    const filteredData = data.filter(row => row[1] !== undefined && row[1] !== null && row[1] !== '');

    // Создание нового листа с отфильтрованными данными
    const newWorksheet = XLSX.utils.aoa_to_sheet(filteredData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, firstSheetName);

    // Сохранение нового файла
    XLSX.writeFile(newWorkbook, outputFilePath);
}

// Использование функции
cleanAndSaveExcel('./oneSheet.xlsx', './deleteEmptyCell.xlsx');
