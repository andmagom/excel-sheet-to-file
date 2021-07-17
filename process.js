const ExcelJS = require('exceljs');

const sen = ['SENNOVA', 'SENN', '8ENNOVA', 'SENNO'];
const bil = ['BILINGUISMO'];
const agro = ['AGROSENA'];
const exceptions = ['Instrucciones', 'Verificación de requisitios'];

async function process3(workbook, id, ids) {
  for (const key of ids) {
    if (key != id) {
      workbook.removeWorksheet(key);
    }
  }
  nameSheet = ""
  cedula = ""
  personName = ""
  idp = ""
  category = ""
  workbook.worksheets.forEach(element => {
    nameSheet = element.name;
    row = element.getRow(10)
    cedula = row.getCell(5).toString().trim()

    row = element.getRow(9)
    personName = row.getCell(5).toString().trim()

    row = element.getRow(16)
    idp = row.getCell(5).toString().trim()

    category = nameSheet.trim().split(' ');
    category = category.pop();
    category = getNameCategory(nameSheet, category);
  });
  if(exceptions.includes(nameSheet)) {
    console.log('Bypass ' + nameSheet + ' Sheet');
    return
  }
  nameOutput = cedula + ' F-230 ' + personName + ' IDP ' + idp + ' ' + category + '.xlsx';
  await workbook.xlsx.writeFile('./result/' + nameOutput);
  console.log('Ready ' + nameOutput);
}

function getNameCategory(nameSheet, name) {
  if( nameSheet.length == 2 || nameSheet.length == 3) {
    if (nameSheet.toUpperCase().startsWith('S')) {
      return 'SENNOVA';
    }
  } else if(nameSheet.includes('SENNOVA')) {
    return 'SENNOVA';
  } else if(nameSheet.includes('SEN')) {
    return 'SENNOVA';
  } else if(nameSheet.includes('BILINGUISMO')) {
    return 'BILINGUISMO';
  } else if (nameSheet.includes('AGROSENA')) {
    return 'AGROSENA';
  } else if (sen.includes(name)) {
    return 'SENNOVA';
  } else if (bil.includes(name)) {
    return 'BILINGUISMO';
  } else if (agro.includes(name)) {
    return 'AGROSENA';
  }
  return name;
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