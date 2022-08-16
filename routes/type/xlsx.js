const express = require('express');
const router = express.Router();

const fs = require('fs')
const path = require('path')
const formidable = require('formidable')
const ExcelJS = require('exceljs')
const logger = require('npmlog')
const moment = require('moment')

moment.locale('ru')

const resault_folder = 'uploads'
const reserved_data_keys = ['file_name']

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        let data_json, template_buffer, data_path, number_start, date_start, date_end = null
        let number_min = 1
        let number_max = 999

        new formidable.IncomingForm().parse(req)
            .on('field', (name, field) => {
                // logger.info('XLSX', `Field name: ${name}, Field value: ${field}`)
                switch (name) {
                    case 'number_start':
                        number_start = field?field:1
                        break;
                    case 'date_start':
                        date_start = new Date(field)
                        break;
                    case 'date_end':
                        date_end = new Date(field)
                        break;
                }
            })
            .on('file', (name, file) => {
                switch (name) {
                    case 'template_file':
                        template_buffer = fs.readFileSync(file.filepath)
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
                    data_json = JSON.parse(fs.readFileSync(data_path))

                    let number_loop = number_start - 1
                    let date_loop = date_start
                    let file_saved_index = 0
                    while (date_loop <= date_end) {
                        let file_save_path = generateFileFolders(date_loop)

                        Array.from(data_json).map( (data, index) => {
                            // NUMBER LOOP START
                            if ( number_loop >= number_max ) {
                                number_loop = number_min - 1
                            }
                            data.index_1 = ++number_loop

                            if ( number_loop >= number_max ) {
                                number_loop = number_min - 1
                            }
                            data.index_2 = data.postfix?data.index_1:++number_loop

                            data.index_1 = data.index_1.toString().padStart(4, "0");
                            data.index_2 = data.index_2.toString().padStart(4, "0");
                            
                            // NUMBER LOOP END
                            
                            let date = moment(date_loop).format('LL').split(' ')
                            data.day = date[0]
                            data.month = date[1]
                            data.year = date[2]

                            let file_name = `${file_saved_index} - ${moment(date_loop).format('L')} (${data.index_1}-${data.index_2}) ` + data.file_name

                            // Create file .xlsx
                            renderWorksheet(file_save_path, file_name, template_buffer, data)
                            
                            file_saved_index++
                        })

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

function generateFileFolders(dateLoop) {
    // current day folder
    if ( !fs.existsSync(resault_folder) ) {
        fs.mkdirSync(resault_folder)
    }
    let res_folder = path.join(resault_folder, `${moment(dateLoop).format('L')}`)
    if ( !fs.existsSync(res_folder) ) {
        fs.mkdirSync(res_folder)
    }

    return res_folder
}

async function renderWorksheet(fileSavePath, fileName, buffer, data) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    var worksheet = workbook.worksheets[0]; 
    let count = 0
    Object.keys(data).map(data_key => {
        if ( !reserved_data_keys.includes(data_key) ) {
            worksheet.eachRow(function(row, rowNumber) {
                row.eachCell(function(cell, colNumber) {
                    if (cell.value && cell.value.formula) {
                        cell.value = { formula: cell.value.formula, result: cell.value.result.replaceAll(`#${data_key}#`, data[data_key]) }
                    }
                    if (cell.value && typeof cell.value == 'string') {
                        cell.value = cell.value.replaceAll(`#${data_key}#`, data[data_key]);
                    }
                });
            });
        }
    })

    fileSavePath = path.join(fileSavePath, fileName+'.xlsx')
    const buffer_res = await workbook.xlsx.writeBuffer();
    return fs.writeFileSync(fileSavePath, buffer_res);
}

module.exports = router;