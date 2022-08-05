const express = require('express');
const router = express.Router();

const fs = require('fs')
const formidable = require('formidable')
const XLSX = require('xlsx')

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        let template_path, data_path = null
        new formidable.IncomingForm().parse(req)
            .on('field', (name, field) => {
                console.log('Field', name, field)
            })
            .on('file', (name, file) => {
                switch (name) {
                    case 'template_file':
                        console.log(`template_file - PATH "${file.filepath}" and MIMETYPE "${file.mimetype}"`, )
                        template_path = file.filepath
                        break;
                    case 'data_file':
                        console.log(`data_file - PATH "${file.filepath}" and MIMETYPE "${file.mimetype}"`, )
                        data_path = file.filepath
                        break;
                }
            })
            .on('aborted', () => {
                console.error('Request aborted by the user')
            })
            .on('error', (err) => {
                console.error('Error', err)
                throw err
            })
            .on('end', () => {
                let data_json = JSON.parse(fs.readFileSync(data_path))
                let buffer = renderWorksheet(template_path, data_json)

                res.statusCode = 200;
                res.setHeader('Content-Disposition', 'attachment; filename="result.xlsx"');
                res.setHeader('Content-Type', 'application/vnd.ms-excel');
                res.end(buffer);
            })
    })

var default_workbook, default_worksheet, default_json = null
function renderWorksheet(templatePath, data) {
    if ( default_json === null ) {
        default_workbook = XLSX.readFileSync(templatePath)
        default_worksheet = default_workbook.Sheets[default_workbook.SheetNames[0]]
        default_json = XLSX.utils.sheet_to_json(default_worksheet)
    }
    
    let updated_json = generateWorksheetJson(default_json, data)

    let new_workbook = XLSX.utils.book_new();
    let updated_worksheet = XLSX.utils.json_to_sheet(updated_json)
    XLSX.utils.book_append_sheet(new_workbook, updated_worksheet, "Page");
    
    return XLSX.write(new_workbook, { type:"buffer", bookType:"xlsx" });
}

function generateWorksheetJson(json, data = null) {
    let string_from_json = JSON.stringify(json)

    let index = '001'
    let day = '2'
    let month = 'августа'
    let year = '2022'
    let number_plate = 'AS-2312'
    let number_garage = '4Д'
    string_from_json = string_from_json.replaceAll('#index#', index)
    string_from_json = string_from_json.replaceAll('#day#', day)
    string_from_json = string_from_json.replaceAll('#month#', month)
    string_from_json = string_from_json.replaceAll('#year#', year)
    string_from_json = string_from_json.replaceAll('#number_plate#', number_plate)
    string_from_json = string_from_json.replaceAll('#number_garage#', number_garage)

    return JSON.parse(string_from_json)
}

module.exports = router;