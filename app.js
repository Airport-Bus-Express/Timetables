const { convertExcelToJson } = require('./xlsx-to-json.js');

const excelPath = process.argv[2]
const colsOutbound = process.argv[3] || "2,6"
const colsReturn = process.argv[4] || "9,12"

if (!excelPath) {
  console.log("ERROR: Please provide the path to the Excel file!");
  process.exit()
}
convertExcelToJson(excelPath, colsOutbound, colsReturn)
  .then(json => {
    if (json) {
      // console.log(JSON.stringify(json, null, 2));
    }
  })
  .catch(error => console.error(error));