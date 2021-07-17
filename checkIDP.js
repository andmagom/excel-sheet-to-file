const ExcelJS = require('exceljs');
const fs = require('fs');

const workbook = new ExcelJS.Workbook();

const manualReview = []
const errors = []

const FOLDER = 'result'

async function processFile(dir, nameFile) {
  workbook2 = await workbook.xlsx.readFile(dir + '/' + nameFile);
  sheetName = ""
  cellValue = ""
  workbook2.worksheets.forEach(element => {
    sheetName = element.name
    row = element.getRow(16)
    cellValue = row.getCell(5).toString()
    cellValue = cellValue.trim()
  });
  sheetIDP = sheetName.split('IDP')
  if (sheetIDP.length != 2) {
    manualReview.push(nameFile)
    return;
  }
  sheetIDP = sheetIDP[1]
  sheetIDP = sheetIDP.trim()
  sheetIDP = sheetIDP.split(' ')
  if (sheetIDP.length != 2) {
    manualReview.push(nameFile)
    return;
  }
  sheetIDP = sheetIDP[0]
  if (sheetIDP != cellValue) {
    console.log('Error detectado en ' + nameFile);
    errors.push(nameFile)
  }
}

function executeOnFiles(dir) {
  fs.readdir(dir, (err, files) => {
    if (err) console.log(err);
    else {
      const promises = [];
      files.forEach((elem) => {
        if (elem.endsWith('.xlsx')) {
          promises.push(processFile(dir, elem));
        }
      });
      Promise.all(promises)
        .then(() => createFileResult())
        .catch(err => console.log(err));
    }
  })
}

function createFileResult() {
  let data = "Errores: \n";
  for (const e of errors) {
    data += e + '\n';
  }
  data += "Revisar Manualmente: \n";
  for (const e of manualReview) {
    data += e + '\n';
  }
  fs.writeFileSync('result.txt', data);
  console.log('FINISHED')
}

executeOnFiles(FOLDER)

//main('20666968 (19 hojas) copy.xlsx')

