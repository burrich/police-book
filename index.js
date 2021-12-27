const { readdirSync } = require('fs');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const XLSX = require('xlsx');

function readPDF(pdfPath) {
  const loadingTask = pdfjsLib.getDocument(pdfPath);
  let data;
  
  return loadingTask.promise
    .then(doc => {
      // Load page
      return doc.getPage(1).then(page => {
        return page
          .getTextContent()
          .then(content => {
            const strings = content.items.map(item => item.str);

            // Remove entries below "N° SIREN/SIRET :"
            strings.splice(0, (strings.lastIndexOf("N° SIREN/SIRET :") + 1));
            // Remove entries above "Date et heure de l'achat..."
            strings.splice(strings.findIndex((str) => str.includes("Date et heure de l'achat")) + 1);

            // Get pdf entryDate, entryTime and entryDateISO 
            const entryDate = strings[strings.length - 1].substring(27, 37);
            const entryTime = strings[strings.length - 1].substring(40, 45);
            const entryDateSplit = entryDate.split('/');
            const entryTimeSplit = entryTime.split('h');
            const entryDateISO = new Date(
              entryDateSplit[2], 
              entryDateSplit[1],
              entryDateSplit[0],
              entryTimeSplit[0],
              entryTimeSplit[1]
            );

            const isLegalPerson = (strings[0] === "Adresse :");
            const siren = isLegalPerson ? strings[8] : null;

            let name = isLegalPerson ? strings[9] : strings[0];
            if (!isLegalPerson && strings[1] != "Adresse :") { 
              name += " " + strings[1] 
            };

            const immat = strings[strings.indexOf("Numéro d'immatriculation") + 2];
            const vin = strings[strings.length - 2];

            // Set pdf data object
            data = {
              entryDate: entryDate,
              entryTime: entryTime,
              entryDateISO : entryDateISO,
              name: name,
              siren: siren,
              immat: immat,
              vin: vin
            };

            // Release page resources.
            page.cleanup();
          })
      });
    })
    .then(
      () => {
        console.log(pdfPath + ' read, data extracted');

        return Promise.resolve(data);
      },
      err => {
        console.error("Error: " + err);
      }
    );
}

function writeXLSX(filename, data) {
  // console.log(new Date());

  const wb = XLSX.utils.book_new(); 
  
  // Add worksheets to workbook with data
  data.forEach((purchases, idx) => {
    const purchasesPropAlias = purchases.map(purchase => {
      return {
        "DATE D'ENTRÉE": purchase.entryDate,
        "HEURE": purchase.entryTime,
        "NOM / RAISON SOCIALE": purchase.name,
        "SIREN": purchase.siren,
        "IMMATRICULATION": purchase.immat,
        "VIN": purchase.vin,
        "NUMÉRO D'ORDRE": purchase.num
      }
    });

    const ws = XLSX.utils.json_to_sheet(
      purchasesPropAlias,
      { header : [ // Custom order
        "NUMÉRO D'ORDRE", 
        "DATE D'ENTRÉE", 
        "HEURE", 
        "NOM / RAISON SOCIALE", 
        "SIREN", 
        "IMMATRICULATION", 
        "VIN"]} 
    );

    // Set read-only and password
    // TODO: random password
    ws['!protect'] = { 
      password: 'admin'
      // formatCells: false,
      // formatColumns: false,
      // formatRows: false,
      // insertColumns: false,
      // insertRows: false,
      // insertHyperlinks: false,
      // deleteColumns: false,
      // deleteRows: false,
      // sort: false,
      // autoFilter: false,
      // pivotTables: false
    }; 

    ws['!cols'] = [
      {wch: 16},
      {wch: 13},
      {wch: 6},
      {wch: 36},
      {wch: 10},
      {wch: 17},
      {wch: 20}
    ];
    
    const wsName = new String(idx + 1);
    XLSX.utils.book_append_sheet(wb, ws, wsName);
  });

  console.log('\n=> Months sheets appended to year workbook ' + filename);

  // Write workbook

  // console.log(new Date());
  XLSX.writeFile(wb, filename);
  // console.log(new Date());

  console.log('=> ' + filename + ' written !\n')
}

function yearPDFsToXLSX(year) {
  const yearPath = sivPath + '/' + year;

  try {
    const yearPurchases = [];
    const monthsPromises = [];
  
    const yearMonths = readdirSync(yearPath); 
    yearMonths.forEach(month => {
  
      const monthPath = yearPath + '/' + month;
      const monthPdfs = readdirSync(monthPath);
  
      let pdfsPromises = [];
      monthPdfs.forEach(pdfName => {
        pdfsPromises.push(readPDF(monthPath + '/' + pdfName));
      });
      
      const monthPromise = Promise.all(pdfsPromises);
      monthsPromises.push(monthPromise);
      monthPromise.then(monthPurchases => {
        // All month PDFs had bean read and data extracted
        // Sort by date
        monthPurchases.sort((a, b) => {
          return a.entryDateISO - b.entryDateISO;
        });
      
        yearPurchases.push(monthPurchases);
      });
    });
    
    // Write xlsx only when all months datas (PDFs) had been processed
    Promise.all(monthsPromises).then(() => {
  
      // Add num property
      yearPurchases.forEach(month => {
        month.forEach(purchase => {
          Object.defineProperty(purchase, 'num', {
            value: fromNum++,
            writable: true, // required ?
            configurable: true, // required ?
            enumerable: true // required ?
          });
        })
      });
  
      // Write
      const policeBookFilename = year + '.xlsx';
      writeXLSX(policeBookFilename, yearPurchases); 
    });
  } catch (err) {
    console.log(err);
  }
};

// MAIN
const sivPath = 'C:/Users/nidet/Google Drive/OAD/SIV';
let fromNum = 1;

console.log('Init...\n');
yearPDFsToXLSX('19');
yearPDFsToXLSX('20');
yearPDFsToXLSX('21');
