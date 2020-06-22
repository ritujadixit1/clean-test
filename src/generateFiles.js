const generateFiles = () => {

    const readXlsxFile = require('read-excel-file/node');
    const xl = require('excel4node');
    const fs = require('fs');
    const constants = require('./constants')

    let mappingData = fs.readFileSync(constants.constants.PATHS.MAPPING_FILE_PATH);
    mappingData = JSON.parse(mappingData);

    const ALPHABETS = 'abcdefghijklmnopqrstuvwxyz';
    let headersMap = new Map()

    readXlsxFile(constants.constants.PATHS.SAMPLE_TEMPLATE_FILE).then((rows) => {
        for(let i = 0; i < rows.length; i++) {
            for(let j = 0; j < rows[0].length; j++) {
                headersMap.set(rows[i][j], [i+1, j+1])
            }
        }
        
        readXlsxFile(constants.constants.PATHS.SAMPLE_DATA_FILE).then((rows) => {

            for(let i = 1; i < rows.length; i++) {
        
                let workbook = new xl.Workbook()
                let worksheet = workbook.addWorksheet('studentData')
        
                for(let columnNames of headersMap.keys()) {
                    worksheet.cell(headersMap.get(columnNames)[0], headersMap.get(columnNames)[1]).string(columnNames)                
                }
        
                for(let j = 0; j < rows[i].length; j++) {
        
                    var cellNumber = mappingData[rows[0][j].toLowerCase()].replace(/\'/g, '').split(/(\d+)/).filter(Boolean)
                    var columnNumber = cellNumber[0]
                    var rowNumber = cellNumber[1]
        
                    var columnIndex = ALPHABETS.search(columnNumber);
                    worksheet.cell(rowNumber, columnIndex + 1).string(rows[i][j].toString())
                }
                
                workbook.write(`output data/mappedData${i}.xlsx`);
            }
        })
    })
}

module.exports = { generateFiles }