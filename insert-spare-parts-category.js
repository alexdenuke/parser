//* После колонки "B" создаётся пустая колонка с названием parts_category и туда вставляются данные в зависимости от названия листа


const XLSX = require('xlsx');

function modifyExcelFile(filePath) {
    // Загрузка файла
    const workbook = XLSX.readFile(filePath);

    // Обработка каждого листа
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        
        // Получаем диапазон данных на листе
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        
        // Для каждой строки...
        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            // Сдвигаем все ячейки на одну колонку вправо начиная с колонки C
            for (let colNum = range.e.c; colNum >= 2; colNum--) {
                const nextCellAddress = XLSX.utils.encode_cell({r: rowNum, c: colNum + 1});
                const cellAddress = XLSX.utils.encode_cell({r: rowNum, c: colNum});
                worksheet[nextCellAddress] = worksheet[cellAddress];
            }
            
            // Очищаем старую колонку C после сдвига
            const newCellAddress = XLSX.utils.encode_cell({r: rowNum, c: 2});
            if (rowNum === 0) {
                worksheet[newCellAddress] = {t: 's', v: 'parts_category'};
            } else {
                worksheet[newCellAddress] = {t: 's', v: sheetName};
            }
        }

        // Обновляем диапазон данных на листе
        range.e.c++;
        worksheet['!ref'] = XLSX.utils.encode_range(range);
    });

    // Сохраняем изменения в новый файл
    XLSX.writeFile(workbook, filePath.replace('.xlsx', 'parts_category.xlsx'));
}

// Запускаем функцию с путем к файлу
modifyExcelFile('./parts.xlsx');
