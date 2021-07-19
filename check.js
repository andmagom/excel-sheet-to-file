const ExcelJS = require('exceljs');
const fs = require('fs');

const RESULT_FOLDER = 'result'

const workbook = new ExcelJS.Workbook();

// CITY
const bogota = ['BOGOTA D.C', 'BOGOTA', 'BOGOTA D.C.', 'BOGOTÀ', 'BOGOTÁ, D.C.', 'BOGOTA  D.C', 'BOGOTA  DC', 'BOGOTA DC'];
const floridablanca = ['FLORIDABLANCA', 'FLORIDA BLANCA'];

// REGIONAL

const valle = ['VALLE', 'VALLE DEL CAUCA'];


async function readResume(nameFile) {
  const workbookResume = await workbook.xlsx.readFile(nameFile);
  workbookResume.getWorksheet('Hoja1').getRow(1).getCell('V').value = 'Validación'
  return workbookResume;
}

async function processFile(resume, folder, nameFile) {
  const workbookCopy = new ExcelJS.Workbook();
  const workbookResult = await workbookCopy.xlsx.readFile(folder + '/' + nameFile);
  cedula = ""
  personName = ""
  idp = ""
  regional = ""
  city = ""

  workbookResult.worksheets.forEach(element => {
    row = element.getRow(10)
    cedula = row.getCell(5).toString().trim()

    row = element.getRow(9)
    personName = row.getCell(5).toString().trim()

    row = element.getRow(16)
    idp = row.getCell(5).toString().trim()

    row = element.getRow(5)
    regional = row.getCell(4).toString().trim()

    row = element.getRow(17)
    city = row.getCell(5).toString().trim()
  });

  checkInResume(resume, cedula, personName, idp, regional, city);
  console.log('Ready ' + nameFile);
}

function validateIDP(real, sheet) {
  real = real.replace(/ /g,'');
  sheet = sheet.replace(/ /g,'');
  return real == sheet;
}

function validateName(real, sheet) {
  real = real.trim();
  sheet = sheet.trim();
  const real2 = real.normalize("NFD").replace(/\p{Diacritic}/gu, "")
  const sheet2 = sheet.normalize("NFD").replace(/\p{Diacritic}/gu, "")
  return real2 == sheet2;
}

function validateRegional(real, sheet) {
  real = real.trim();
  sheet = sheet.trim();
  const real2 = real.normalize("NFD").replace(/\p{Diacritic}/gu, "")
  const sheet2 = sheet.normalize("NFD").replace(/\p{Diacritic}/gu, "")

  if (real2 == sheet2) {
    return true;
  } else if (valle.includes(real2) && valle.includes(sheet2)) {
    return true;
  }

  return real2 == sheet2;

  /*else if (atlantico.includes(real) && atlantico.includes(sheet)) {
    return true;
  } else if(quindio.includes(real) && quindio.includes(sheet)) {
    return true;
  } else if(bolivar.includes(real) && bolivar.includes(sheet)) {
    return true;
  }
  return real == sheet;
  */
}

function validateCity(real, sheet) {
  real = real.trim().toUpperCase();
  sheet = sheet.trim().toUpperCase();
  const real2 = real.normalize("NFD").replace(/\p{Diacritic}/gu, "")
  const sheet2 = sheet.normalize("NFD").replace(/\p{Diacritic}/gu, "")

  if (real2 == sheet2) {
    return true;
  } else if (bogota.includes(real2) && bogota.includes(sheet2)) {
    return true;
  } else if (floridablanca.includes(real2) && floridablanca.includes(sheet2)) {
    return true;
  }

  return real2 == sheet2;

  /*else if(medellin.includes(real) && medellin.includes(sheet)) {
    return true;
  } else if(puertoasis.includes(real) && puertoasis.includes(sheet)) {
    return true;
  } else if(ibague.includes(real) && ibague.includes(sheet)) {
    return true;
  } else if(garzon.includes(real) && garzon.includes(sheet)) {
    return true;
  } else if(giron.includes(real) && giron.includes(sheet)) {
    return true;
  } else if(apartado.includes(real) && apartado.includes(sheet)) {
    return true;
  } else if(sanandrestumaco.includes(real) && sanandrestumaco.includes(sheet)) {
    return true;
  } else if(itagui.includes(real) && itagui.includes(sheet)) {
    return true;
  } else if(popayan.includes(real) && popayan.includes(sheet)) {
    return true;
  } else if(malaga.includes(real) && malaga.includes(sheet)) {
    return true;
  }
  */
  return real == sheet;
}

function checkInResume(resume, cedula, personName, idp, regional, city) {
  const workSheet = resume.getWorksheet('Hoja1');
  const column = workSheet.getColumn(2);
  const cedulas = column.values;
  for (i = 0; i < cedulas.length; i++) {
    const ced = cedulas[i];
    if (ced == cedula) {
      const row = workSheet.getRow(i);
      const personNameRow = row.getCell('C').toString();
      const idpRow = row.getCell('F').toString();
      const regionalRow = row.getCell('H').toString();
      const cityRow = row.getCell('J').toString();

      const nameB = validateName(personNameRow, personName);
      const idpB = validateIDP(idpRow, idp);
      const regionalB = validateRegional(regionalRow, regional);
      const cityB = validateCity(cityRow, city);

      if (nameB && idpB && regionalB && cityB) {
        workSheet.getRow(i).getCell('V').value = 'validado';
        break;
      }
    }
  }
}

async function executeOnFiles(dir, resume) {
  files = fs.readdirSync(dir);
  for (const elem of files) {
    if (elem.endsWith('.xlsx')) {
      await processFile(resume, RESULT_FOLDER, elem);
    }
  }
  await workbook.xlsx.writeFile('validación.xlsx');
  console.log('FINISHED');
}

function main(nameFile) {
  readResume(nameFile)
    .then(resume => executeOnFiles(RESULT_FOLDER, resume))
    .catch(err => console.log(err))
}

main('BASE DE DATOS REGIONAL CALDAS.xlsx')