const { convertExcelToJson } = require('./xlsx-to-json.js');

const excelPath = process.argv[2]

if (!excelPath) {
  console.log("ERROR: Please provide the path to the Excel file!");
  process.exit()
}
convertExcelToJson(excelPath)
  .then(json => {
    if (json) {
      // console.log(JSON.stringify(json, null, 2));
    }
  })
  .catch(error => console.error(error));