const express = require('express');
const router = express.Router();

const fs = require('fs')
const AdmZip = require("adm-zip");
const path = require('path')
const formidable = require('formidable')
const ExcelJS = require('exceljs')
const logger = require('npmlog')
const moment = require('moment')
const nodemailer = require("nodemailer")

require('dotenv').config()
moment.locale('ru')

const resault_folder = 'uploads'
const reserved_data_keys = ['file_name']

const zip = new AdmZip();
var zip_path;

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        let data_json, template_buffer, data_path, number_start, date_start, date_end, email = null
        let number_min = 1
        let number_max = 999

        new formidable.IncomingForm().parse(req)
            .on('field', (name, field) => {
                // logger.info('XLSX', `Field name: ${name}, Field value: ${field}`)
                switch (name) {
                    case 'number_start':
                        number_start = field ? field : 1
                        break;
                    case 'date_start':
                        date_start = new Date(field)
                        break;
                    case 'date_end':
                        date_end = new Date(field)
                        break;
                    case 'email':
                        email = field
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

                    zip_path = path.join(resault_folder, `${moment(Date.now()).format('L')}.zip`)
                    zip.writeZip(zip_path);

                    let number_loop = number_start - 1
                    let date_loop = date_start
                    let file_saved_index = 0
                    while (date_loop <= date_end) {
                        Array.from(data_json).map((data, index) => {
                            // NUMBER LOOP START
                            if (number_loop >= number_max) {
                                number_loop = number_min - 1
                            }
                            data.index_1 = ++number_loop

                            if (number_loop >= number_max) {
                                number_loop = number_min - 1
                            }
                            data.index_2 = data.postfix ? data.index_1 : ++number_loop

                            data.index_1 = data.index_1.toString().padStart(4, "0");
                            data.index_2 = data.index_2.toString().padStart(4, "0");

                            // NUMBER LOOP END

                            let date = moment(date_loop).format('LL').split(' ')
                            data.day = date[0]
                            data.month = date[1]
                            data.year = date[2]

                            let file_name = `${file_saved_index} - ${moment(date_loop).format('L')} (${data.index_1}-${data.index_2}) ` + data.file_name

                            // Create file .xlsx
                            renderWorksheet(file_name, template_buffer, data, date_loop)

                            file_saved_index++
                        })

                        logger.info('XLSX', 'Loading ...')

                        date_loop = new Date(date_loop.setDate(date_loop.getDate() + 1));
                    }

                    sendFileToEmail(email).catch(console.error);
                    logger.info('XLSX', 'Processed successfully!')
                    res.statusCode = 200;
                    res.render('success', {
                        title: 'Обработка выполнена успешно',
                        email: email,
                        download_file_path: path.resolve(zip_path)
                    });
                } catch (error) {
                    logger.error('XLSX', 'Processed with errors: %j', error)
                    res.statusCode = 500;
                    res.end('Processed with errors!');
                }
            })
    })

async function updateZipArchive(filePath, buffer) {
    try {
        zip.addFile(filePath, buffer);
        zip.writeZip(zip_path);
    } catch (error) {
        logger.error('ZIP - UPDATE', 'Processed with errors: %j', error)
    }
}

function renderWorksheet(fileName, buffer, data, dateLoop) {
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.xlsx.load(buffer).then(res => {
            var worksheet = workbook.worksheets[0]; 
            let count = 0
            Object.keys(data).map(data_key => {
                if (!reserved_data_keys.includes(data_key)) {
                    worksheet.eachRow(function (row, rowNumber) {
                        row.eachCell(function (cell, colNumber) {
                            let replace = `#${data_key}#`;
                            let replace_regex = new RegExp(replace, 'g')
                            if (cell.value && cell.value.formula) {
                                let curretValue = cell.value.result.toString()
                                cell.value = { formula: cell.value.formula, result: curretValue.replace(replace_regex, data[data_key]) }
                            }
                            if (cell.value && typeof cell.value == 'string') {
                                let curretValue = cell.value.toString()
                                cell.value = curretValue.replace(replace_regex, data[data_key]);
                            }
                        });
                    });
                }
            })

            let file_path = path.join(`${moment(dateLoop).format('L')}`, fileName + '.xlsx')
            workbook.xlsx.writeBuffer().then(buffer_res => {
                updateZipArchive(file_path, buffer_res)
            })
        })
    } catch (error) {
        logger.error('ExcelJS - UPDATE', 'Processed with errors: %j', error)
    }
}

async function sendFileToEmail(email) {
    let mailConfig;
    
    mailConfig = {
        host: process.env.SMTP_HOST,
        port: process.env.SMTP_PORT, 
        secure: true,
        authMethod: 'LOGIN',
        auth: {
            user: process.env.SMTP_EMAIL,
            pass: process.env.SMTP_PASSWORD,
        },
        attachments: [  
            {
                path: path.resolve(zip_path),
                contentType: 'application/zip'
            }
        ]
    };

    let transporter = nodemailer.createTransport(mailConfig);

    let info = await transporter.sendMail({
        from: process.env.SMTP_EMAIL,
        to: email,
        subject: "BulkEditor App - Files",
        html: "<h3>Спасибо что воспользовались BulkEditor App!</h3>",
    });

    console.log("Message sent: %s", info.messageId);
}

module.exports = router;