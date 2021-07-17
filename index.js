const ExcelJS = require('exceljs');
const fs = require('fs');

const process = require('./process');
const workbook = new ExcelJS.Workbook();

const RESULT_FOLDER = 'result'

async function processFile(nameFile) {
  await workbook.xlsx.readFile(nameFile)
    .then(process.getTotalSheets)
    .then(res => loop(res.workbook, nameFile))
    .catch(err => console.log(err));
}

async function loop(workbook, nameFile) {
  const ids = [];
  workbook.worksheets.forEach(element => {
    ids.push(element.id);
  });
  for (const key of ids) {
    workbookCopy = await workbook.xlsx.readFile(nameFile);
    await process.process3(workbookCopy, key, ids)
  }
}

function createFolder(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir)
  }
}

function executeOnFiles(dir) {
  fs.readdir(dir, (err, files) => {
    if (err) console.log(err);
    else {
      const promises = [];
      files.forEach(elem => {
        if (elem.endsWith('.xlsx')) {
          promises.push(processFile(elem));
        }
      });
      Promise.all(promises)
        .then(() => console.log('FINISHED'))
        .catch(err => console.log(err));
    }
  })
}

function main() {
  createFolder(RESULT_FOLDER);
  executeOnFiles('.');
}

main();
