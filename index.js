//* Указать папку с моделями оборудования. Скрипт анализирует структуру этой папки и создаёт на основе неё файл эксель. 1 колонка это название модели, 2 колонка это категория модели, 3 колонка это бренд модели. 

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

async function scanDirectories(basePath) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Products');

    // Добавляем заголовки для колонок
    sheet.columns = [
        { header: 'Model', key: 'model', width: 30 },
        { header: 'Category', key: 'category', width: 30 },
        { header: 'Brand', key: 'brand', width: 30 }
    ];

    // Рекурсивная функция для обхода директорий
    async function readDir(dir, brand, category) {
        const items = fs.readdirSync(dir, { withFileTypes: true });
        for (let item of items) {
            const fullPath = path.join(dir, item.name);
            if (item.isDirectory()) {
                if (!brand) {
                    await readDir(fullPath, item.name);
                } else if (!category) {
                    await readDir(fullPath, brand, item.name);
                } else {
                    await readDir(fullPath, brand, category);
                    sheet.addRow({ model: item.name, category: category, brand: brand });
                }
            }
        }
    }

    await readDir(basePath);

    // Записываем файл
    await workbook.xlsx.writeFile('models.xlsx');
    console.log('Excel file has been written.');
}

// Запуск с путем к корневой папке
scanDirectories('./equipment');
