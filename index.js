const fs = require('fs');
const XLSX = require('xlsx');
const prompt = require('prompt');
prompt.start();

const arrayToObjectReverse = arr => {
  return arr.reduce((curObj, element, i) => {
    curObj[element] = i + 1;
    return curObj;
  }, {})
}


const getExcelArr = filePath => {
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const worksheet = wb.Sheets[sheetName];
  const worksheetJSON = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  const excelArr = worksheetJSON.map(cell => cell[0]);
  return excelArr;
};

const createResExcelFile = resArr => {
  const resWorksheet = XLSX.utils.json_to_sheet(resArr);

  const resWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(resWorkbook, resWorksheet, 'result');
  XLSX.writeFile(resWorkbook, 'res.xlsx');

  console.log("created result file successfully");
  console.log('file name is res.xlsx');
}

const main = async () => {
  const { filePath, folderPath } = await prompt.get(['filePath', 'folderPath']);
  const dirContent = arrayToObjectReverse(fs.readdirSync(folderPath));
  const excelArr = getExcelArr(filePath);

  const resArr = excelArr.map(cell => ({ cell, isExist: dirContent[cell] ? true : false }));

  createResExcelFile(resArr);
}

main();