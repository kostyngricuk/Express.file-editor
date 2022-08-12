const express = require('express');
const router = express.Router();

const fs = require('fs')
const path = require('path')
const formidable = require('formidable')
const XLSX = require('xlsx')
const logger = require('npmlog')
const moment = require('moment')

moment.locale('ru')

const resault_folder = 'uploads'
const reserved_data_keys = ['file_name']

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        let data_json, template_path, data_path, number_start, date_start, date_end = null
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
                            // NUMBER LOOP END
                            
                            let date = moment(date_loop).format('LL').split(' ')
                            data.day = date[0]
                            data.month = date[1]
                            data.year = date[2]

                            let file_name = `${file_saved_index} - ${moment(date_loop).format('L')} (${data.index_1}-${data.index_2}) ` + data.file_name

                            // Create file .xlsx
                            renderWorksheet(file_save_path, file_name, template_path, data)
                            
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

function renderWorksheet(fileSavePath, fileName, templatePath, data) {
    fileSavePath = path.join(fileSavePath, fileName+'.xlsx')

    let wb = XLSX.readFileSync(templatePath, {
        cellNF: true,
        cellStyles: true
    })
    let ws = wb.Sheets[wb.SheetNames[0]]
    let ref = XLSX.utils.decode_range(ws["!ref"]);

    Object.keys(data).map(data_key => {
        if ( !reserved_data_keys.includes(data_key) ) {
            for (var R = ref.s.r; R <= ref.e.r; R++) {
                for (var C = ref.s.c; C <= ref.e.c; C++) {
                    let cell_name = XLSX.utils.encode_cell({ c: C, r: R });
                    if (ws[cell_name] && ws[cell_name].t == 's') {
                        Object.keys(ws[cell_name]).forEach(key => {
                            let current_value = ws[cell_name][key].toString()
                            ws[cell_name][key] = current_value.replace(`#${data_key}#`, data[data_key]);
                        });
                    }
                }
            }
        }
    })

    ws["!ref"] = XLSX.utils.encode_range(ref);

    return XLSX.writeFile(wb, fileSavePath, {
        themeXLSX: true, 
        compression: true
    })
}

module.exports = router;