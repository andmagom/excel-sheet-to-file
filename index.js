const ExcelJS = require('exceljs');
const fs = require('fs');

const process = require('./process');
const workbook = new ExcelJS.Workbook();

const RESULT_FOLDER = 'result'
const ORIGINAL_FOLDER = 'excels'

async function processFile(folder, nameFile) {
  await workbook.xlsx.readFile(folder + '/' + nameFile)
    .then(res => loop(res, folder, nameFile))
    .catch(err => console.log(err));
}

async function loop(workbook, folder, nameFile) {
  const ids = [];
  workbook.worksheets.forEach(element => {
    ids.push(element.id);
  });
  for (const key of ids) {
    const workbook2 = new ExcelJS.Workbook();
    workbookCopy = await workbook2.xlsx.readFile(folder + '/' + nameFile);
    await process.process3(workbookCopy, key, ids)
  }
}

function createFolder(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir)
  }
}

async function executeOnFiles(dir) {
  files = fs.readdirSync(dir);
  for (const elem of files) {
    if (elem.endsWith('.xlsx')) {
      await processFile(ORIGINAL_FOLDER, elem);
    }
  }
}

function main() {
  createFolder(RESULT_FOLDER);
  executeOnFiles(ORIGINAL_FOLDER)
    .then(() => console.log('FINISHED'))
    .catch(err => console.log(err));
}

main();
