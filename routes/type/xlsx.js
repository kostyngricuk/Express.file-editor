const express = require('express');
const router = express.Router();

const fs = require('fs')
const path = require('path')
const formidable = require('formidable')
const XLSX = require('xlsx')
const logger = require('npmlog');

const resault_folder = 'uploads'

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        let template_path, data_path, number_start, number_end, date_start, date_end = null
        let use_strict = false
        new formidable.IncomingForm().parse(req)
            .on('field', (name, field) => {
                // logger.info('XLSX', `Field name: ${name}, Field value: ${field}`)
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
                logger.warn('XLSX', 'Request aborted by the user')
            })
            .on('error', (err) => {
                logger.error('XLSX', 'Error: %j', err)
                throw err
            })
            .on('end', () => {
                try {
                    let data_json = JSON.parse(fs.readFileSync(data_path))

                    let date_loop = date_start
                    while (date_loop <= date_end) {
                        let file_save_path = generateFileFolders(date_loop, use_strict)
                        let file_name = 'file_name'

                        if (!use_strict) {
                            file_name += '-' + date_loop.toDateString().replaceAll(' ', '_')
                        }
                        // ...
                        renderWorksheet(file_save_path, file_name, template_path, data_json)

                        logger.info('XLSX', 'Loading ...')

                        date_loop = new Date(date_loop.setDate(date_loop.getDate() + 1));
                    }

                    logger.info('XLSX', 'Processed successfully!')
                    res.statusCode = 200;
                    res.end('Processed successfully!');
                } catch (error) {
                    logger.error('XLSX', 'Processed with errors: %j', error)
                    res.statusCode = 500;
                    res.end('Processed with errors!');
                }
            })
    })

function generateFileFolders(dateStart, useStrict = false) {
    // current day folder
    let res_folder = path.join(resault_folder, new Date().toDateString().replaceAll(' ', '_'))
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

function renderWorksheet(fileSavePath, fileName, templatePath, data) {
    fileSavePath = path.join(fileSavePath, fileName+'.xlsx')

    let wb = XLSX.readFileSync(templatePath, {
        cellNF: true,
        cellStyles: true
    })
    let ws = wb.Sheets[wb.SheetNames[0]]
    let ref = XLSX.utils.decode_range(ws["!ref"]);

    for (var R = ref.s.r; R <= ref.e.r; R++) {
        for (var C = ref.s.c; C <= ref.e.c; C++) {
            let cell_name = XLSX.utils.encode_cell({ c: C, r: R });
            if (ws[cell_name] && ws[cell_name].t == 's') {
                Object.keys(ws[cell_name]).forEach(key => {
                    let current_value = ws[cell_name][key].toString()
                    ws[cell_name][key] = updateCellValue(current_value);
                });
            } else {
                continue;
            }
        }
    }

    ws["!ref"] = XLSX.utils.encode_range(ref);

    return XLSX.writeFile(wb, fileSavePath, {
        themeXLSX: true, 
        compression: true
    })
}

function updateCellValue(currentValue) {
    let new_value = currentValue.replaceAll('#index#', '001')
    return new_value
}

module.exports = router;