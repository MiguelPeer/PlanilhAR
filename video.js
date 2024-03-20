const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');


function readExcelFiles(folderPath) {
    const files = fs.readdirSync(folderPath);
    const excelFiles = files.filter(file => path.extname(file) === '.xlsx');

    const data = [];
    let isFirstSheet = true; 

    for (const file of excelFiles) {
        const workbook = XLSX.readFile(path.join(folderPath, file));
        const sheetNames = workbook.SheetNames;

       
        let baseSheetName = null;
        for (const sheetName of sheetNames) {
            if (sheetName === "Plan1") {
                baseSheetName = sheetName;
                break;
            }
        }

        if (baseSheetName) {
            const worksheet = workbook.Sheets[baseSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });


            let shippingIndex;
            let dateIndex;
            let differenceIndex;

            const newArray = [];

            for (let  i = 0; i < jsonData.length; i++) {
              const row = jsonData[i];

              if (i === 0) {
                shippingIndex = row.findIndex((columm) => columm === "DATA");
                dateIndex = row.findIndex((columm) => columm === "SITUAÇAO" );
                differenceIndex = row.findIndex((columm) => columm === "VALOR");
                
                newArray.push(["DATA_NEW", "SIT_NEW", "VALOR_NEW"])
                continue;
              }
              
              const shipping = row[shippingIndex];
              const date = row[dateIndex];
              const difference = row[differenceIndex];

              
              newArray.push([shipping, date, difference])

              console.log(shipping, date, difference)

            }

            
            if (isFirstSheet) {
                data.push({ sheetName: baseSheetName, data: newArray });
                isFirstSheet = false;
            } else {
      
                data.push({ sheetName: baseSheetName, data: newArray.slice(1) });
            }
        }
    }
    
    return data;
}


function createConsolidatedExcel(data, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Consolidated Data');

    data.forEach(sheet => {
        sheet.data.forEach(row => {
            worksheet.addRow(row);
        });
    });

    workbook.xlsx.writeFile(outputPath)
        .then(() => {
            console.log('Arquivo Excel consolidado criado com sucesso:', outputPath);
        })
        .catch(error => {
            console.error('Erro ao criar o arquivo Excel consolidado:', error);
        });
}


const folderPath = './entrada';


const outputPath = './saida/video.xlsx';


const excelData = readExcelFiles(folderPath);


createConsolidatedExcel(excelData, outputPath);
