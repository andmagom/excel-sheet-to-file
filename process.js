const ExcelJS = require('exceljs');

async function process3(workbook, id, ids) {
  for (const key of ids) {
    if (key != id) {
      workbook.removeWorksheet(key);
    }
  }
  nameSheet = ""
  workbook.worksheets.forEach(element => {
    nameSheet= element.name;
  });

  await workbook.xlsx.writeFile('./result/' + nameSheet + '.xlsx');
  console.log('Ready ' + nameSheet);
}

function getTotalSheets(workbook) {
  const length = workbook.worksheets.length;
  return {
    length,
    workbook
  };
}

module.exports = {
  process3,
  getTotalSheets
}