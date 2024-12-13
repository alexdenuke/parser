//* Можно указать столбик откуда будут браться данные. Пройтись по этим данным и записать все уникальные значения ( учитывая множественные через запятую) в другой файл эксель 

const xlsx = require('xlsx');
const _ = require('lodash');
const fs = require('fs');

// Путь к входному файлу
const inputFilePath = './xlsx/models.xlsx';
// Путь к выходному файлу
const outputFilePath = './models_categories.xlsx';
// Номер столбца, откуда будут браться данные (начинается с 0)
const columnIndex = 1; // Измените это значение на нужный номер столбца

// Чтение входного файла Excel
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Преобразование листа в JSON
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// Извлечение данных из указанного столбца и обработка уникальных значений
const columnData = data.map(row => row[columnIndex]).filter(Boolean);
const uniqueValues = _.uniq(
    columnData
    .flatMap(cell => cell.split(',').map(value => value.trim()))
);

// Создание нового листа с уникальными значениями
const newSheetData = [['Column Header'], ...uniqueValues.map(value => [value])];
const newSheet = xlsx.utils.aoa_to_sheet(newSheetData);

// Создание новой книги и добавление листа в неё
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Unique Values');

// Запись нового файла Excel
xlsx.writeFile(newWorkbook, outputFilePath);

console.log('New Excel file created successfully.');
