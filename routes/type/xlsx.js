const express = require('express');
const router = express.Router();

const fs = require('fs')
const path = require('path')
const formidable = require('formidable')
const XLSX = require('xlsx')

const resault_folder = 'uploads'

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        let template_path, data_path, number_start, number_end, date_start, date_end = null
        let use_strict = false
        new formidable.IncomingForm().parse(req)
            .on('field', (name, field) => {
                console.log(`Field name: ${name}, Field value: ${field}`)
                switch (name) {
                    case 'number_start':
                        number_start = field?field:1
                        break;
                    case 'number_end':
                        number_end = field?field:999
                        break;
                    case 'date_start':
                        date_start = new Date(field)
                        break;
                    case 'date_end':
                        date_end = new Date(field)
                        break;
                    case 'use_strict':
                        use_strict = true
                        break;
                }
            })
            .on('file', (name, file) => {
                switch (name) {
                    case 'template_file':
                        template_path = file.filepath
                        break;
                    case 'data_file':
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

                while (date_start <= date_end) {
                    let resFilePath = generateFileFolders(date_start, use_strict)

                    console.log(resFilePath)

                    let next_day = date_start.setDate(date_start.getDate() + 1);
                    date_start = new Date(next_day);
                }

                // let buffer = renderWorksheet(template_path, data_json)res.statusCode = 200;
                // res.setHeader('Content-Disposition', 'attachment; filename="result.xlsx"');
                // res.setHeader('Content-Type', 'application/vnd.ms-excel');
                // res.end(buffer);
            })
    })

function generateFileFolders(dateStart, useStrict = false) {
    // current day folder
    let res_folder = path.join(resault_folder, new Date().toDateString())
    if ( !fs.existsSync(res_folder) ) {
        fs.mkdirSync(res_folder)
    }

    // sub folders (strict)
    if ( useStrict ) {
        let date_year = dateStart.getFullYear().toString()
        let date_month = (dateStart.getMonth() + 1).toString()
        let date_day = dateStart.getDate().toString()
        if ( !fs.existsSync(path.join(res_folder, date_year)) ) {
            fs.mkdirSync(path.join(res_folder, date_year));
        }
        if ( !fs.existsSync(path.join(res_folder, date_year, date_month)) ) {
            fs.mkdirSync(path.join(res_folder, date_year, date_month));
        }

        let filePath = path.join(res_folder, date_year, date_month, date_day)
        if ( !fs.existsSync(filePath) ) {
            fs.mkdirSync(filePath);
        }
        return filePath
    }

    return res_folder
}

function renderWorksheet(templatePath, data) {
    let wb = XLSX.readFileSync(templatePath)
    let ws = wb.Sheets[wb.SheetNames[0]]
    let ref = XLSX.utils.decode_range(ws["!ref"]);

    var wscols = new Array();

    for (var r = ref.s.r; r <= ref.e.r; r++) {
        for (var c = ref.s.c; c <= ref.e.c; c++) {
            if (r === ref.s.r) {
                wscols.push({ wpx: 9 })
            }
            
            let cell_name = XLSX.utils.encode_cell({ c: c, r: r });
            if (ws[cell_name]) {
                let current_value = ws[cell_name].v.toString()
                let update_value = updateCellValue(current_value)
                ws[cell_name] = { t: 's', v: update_value };
            } else {
                continue;
            }
        }
    }

    ws['!cols'] = wscols;
    ws["!ref"] = XLSX.utils.encode_range(ref);

    return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function updateCellValue(currentValue) {
    let new_value = currentValue.replaceAll('#index#', '001')
    return new_value
}

module.exports = router;