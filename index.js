const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

// Função para ler os arquivos Excel de uma pasta
function readExcelFiles(folderPath) {
    const files = fs.readdirSync(folderPath);
    const excelFiles = files.filter(file => path.extname(file) === '.xlsx');

    const data = [];
    let isFirstSheet = true; // Variável para rastrear se estamos lidando com a primeira planilha

    for (const file of excelFiles) {
        const workbook = XLSX.readFile(path.join(folderPath, file));
        const sheetNames = workbook.SheetNames;

        // Verifica se há uma planilha com o nome "Base" e pega apenas os dados dessa planilha
        let baseSheetName = null;
        for (const sheetName of sheetNames) {
            if (sheetName === "Base") {
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
                shippingIndex = row.findIndex((columm) => columm === "REMESSA");
                dateIndex = row.findIndex((columm) => columm === "DATA" );
                differenceIndex = row.findIndex((columm) => columm === "DIFERENCA");
                
                newArray.push(["CRE_CODI", "CRE_DATA", "REC_VLR_DIFERENCA", "REC_CD_TITULO", "REC_IMPOSTO_DT_BASE", "REC_FLG_PROCESSADO"])
                continue;
              }
              
              const shipping = row[shippingIndex];
              const date = row[dateIndex];
              const difference = row[differenceIndex];

              
              newArray.push([shipping, date, difference])

              console.log(shipping, date, difference)

            }

            // Se for a primeira planilha, adicione o cabeçalho
            if (isFirstSheet) {
                data.push({ sheetName: baseSheetName, data: newArray });
                isFirstSheet = false;
            } else {
                // Se não for a primeira planilha, adicione apenas os dados, sem o cabeçalho
                data.push({ sheetName: baseSheetName, data: newArray.slice(1) });
            }
        }
    }
    
    return data;
}

// Função para criar um arquivo Excel consolidado
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

// Caminho da pasta com os arquivos Excel
const folderPath = './entrada';

// Caminho de saída para o arquivo Excel consolidado
const outputPath = './saida/DW_RECUPERAR_IMPOSTOS.xlsx';

// Ler os arquivos Excel da pasta
const excelData = readExcelFiles(folderPath);

// Criar o arquivo Excel consolidado
createConsolidatedExcel(excelData, outputPath);
