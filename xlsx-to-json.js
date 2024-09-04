const fs = require('fs');
const ExcelJS = require('exceljs');

const isString = value => typeof value === 'string' || value instanceof String;

function formatExcelTimeAsTime(number) {
  // Split the number into hours and decimal part
  let hours = Math.floor(number);
  let decimalPart = number - hours;

  // Convert decimal part to minutes (multiply by 100, not 60)
  let minutes = Math.round(decimalPart * 100);

  // Handle case where minutes are 60
  if (minutes === 60) {
    hours += 1;
    minutes = 0;
  }

  // Format hours and minutes, ensuring two digits for each
  const formattedHours = hours.toString().padStart(2, '0');
  const formattedMinutes = minutes.toString().padStart(2, '0');

  return `${formattedHours}:${formattedMinutes}`;
}

const tableHeaders = {
  "STD": "Stansted Airport",
  "STR": "Stratford",
  "BETH": "Bethnal Green",
  "LIV": "Liverpool Street",
}

function processCell(cell, isFirst = false) {
  const cellFinalValue = cell?.result || cell
  if (isFirst) {
    return tableHeaders[cellFinalValue]
  } else {
    const cellFormatted = !isNaN(parseFloat(cellFinalValue)) ? formatExcelTimeAsTime(cellFinalValue) : cellFinalValue
    return cellFormatted
  }
}

async function convertExcelToJson(filePath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0];

    // Find the row index where the second table starts
    let secondTableStartRow = 0;
    let firstTable = false;
    let firstTableLastCol = false;
    let secondTable = false;
    let firstCol = 0
    let secondNullCol = -1
    for (let row = 1; row <= worksheet.rowCount; row++) {
      for (let col = 1; col <= worksheet.columnCount; col++) {
        const cell = worksheet.getCell(row, col).value
        if (cell && !firstTable) {
          firstTable = { row: row, col: col };
          firstCol = col
        }
        if (!cell && firstTable && !firstTableLastCol && col > firstCol) {
          firstTableLastCol = col - 1
          firstTable.lastCol = col - 1
        }
        if (!cell && firstTable && secondNullCol === -1) {
          secondNullCol = col
        }
        if (cell && firstTable && !secondTable && col > secondNullCol && secondNullCol > -1) {
          // console.log("row", row, "col", col, secondNullCol)
          secondTable = { row: row, col: col };
        }
        if (cell && secondTable) {
          secondTable.lastCol = col
        }
      }
    }
    console.log("firstTable", firstTable)
    console.log("secondTable", secondTable)

    const firstTableRows = [];
    for (let row = firstTable.row; row < worksheet.rowCount; row++) {
      const rowData = [];
      for (let col = firstTable.col; col <= firstTable.lastCol; col++) {
        const cellValue = worksheet.getCell(row, col).value;
        const cell = processCell(cellValue, firstTableRows.length === 0)
        rowData.push(cell);
      }
      firstTableRows.push(rowData);
    }

    const secondTableRows = [];
    for (let row = secondTable.row; row < worksheet.rowCount; row++) {
      const rowData = [];
      for (let col = secondTable.col; col <= secondTable.lastCol; col++) {
        const cellValue = worksheet.getCell(row, col).value;
        const cell = processCell(cellValue, secondTableRows.length === 0)
        rowData.push(cell);
      }
      secondTableRows.push(rowData);
    }

    // Create the JSON object
    const json = {
      meta: {
        msg: `Converted from Excel file: ${filePath}`,
        filename: filePath,
        date: new Date().toISOString()
      },
      times: {
        firstTable: firstTableRows,
        secondTable: secondTableRows
      }
    };

    const jsonString = JSON.stringify(json, null, 2);
    fs.writeFileSync('dist/latest-timetable.json', jsonString);

    return json;
  } catch (error) {
    console.error('Error:', error);
    return null;
  }
}

module.exports = { convertExcelToJson }