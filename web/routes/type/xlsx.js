const express = require('express');
const router = express.Router();

const fs = require('fs')
const AdmZip = require("adm-zip");
const path = require('path')
const formidable = require('formidable')
const ExcelJS = require('exceljs')
const logger = require('npmlog')
const moment = require('moment')
const nodemailer = require("nodemailer");
const { render } = require('ejs');

require('dotenv').config()
moment.locale('ru')

const resault_folder = 'uploads'
const reserved_data_keys = ['file_name']

const number_min = 1
const number_max = 999
var number_loop = 0

/* GET xlsx */
router.route('/')
    .post((req, res, next) => {
        const zip = new AdmZip();
        const zip_path = path.join(resault_folder, `${moment(Date.now()).format('L')}.zip`)
        
        
        let data_json, template_buffer, data_path, number_start, date_start, date_end, email = null
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
                    
                    number_loop = number_start - 1
                    
                    let promises = []

                    let date_loop = date_start
                    while (date_loop <= date_end) {
                        let date = moment(date_loop)
                        Array.from(data_json).map((data, index) => {
                            promises.push(renderWorksheet(zip, template_buffer, date, data))
                        })

                        logger.info('XLSX', date.format('L') + ' - Loading ...')

                        date_loop = new Date(date_loop.setDate(date_loop.getDate() + 1));
                    }

                    Promise.all(promises).then(() => {
                        zip.writeZip(zip_path);

                        // send files to email
                        sendFileToEmail(zip_path, email).catch(console.error);

                        // response
                        logger.info('XLSX', 'Processed successfully!')
                        res.statusCode = 200;
                        res.render('success', {
                            title: 'Обработка выполнена успешно',
                            email: email,
                            download_file_path: zip_path
                        });
                    })
                } catch (error) {
                    logger.error('XLSX', 'Processed with errors: %j', error)
                    res.statusCode = 500;
                    res.end('Processed with errors!');
                }
            })
    })

function getObjectXLSX(date, data) {
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

    // DATE START
    let dateArray = date.format('LL').split(' ')
    data.day = dateArray[0]
    data.month = dateArray[1]
    data.year = dateArray[2]
    // DATE END

    let obj = {
        fileName: `${date.format('L')} (${data.index_1}-${data.index_2}) ` + data.file_name,
        dataFile: data, 
        folderName: date.format('L')
    }
    return obj
}

async function renderWorksheet(zip, template_buffer, dateLoop, data) {
    let workbook = new ExcelJS.Workbook();
    
    await workbook.xlsx.load(template_buffer)
    var worksheet = workbook.worksheets[0];

    let obj = getObjectXLSX(dateLoop, data)
    Object.keys(obj.dataFile).map(data_key => {
        if (!reserved_data_keys.includes(data_key)) {
            worksheet.eachRow(function (row, rowNumber) {
                row.eachCell(function (cell, colNumber) {
                    let replace = `#${data_key}#`;
                    let replace_regex = new RegExp(replace, 'g')
                    if (cell.value && cell.value.formula) {
                        let curretValue = cell.value.result.toString()
                        cell.value = { formula: cell.value.formula, result: curretValue.replace(replace_regex, obj.dataFile[data_key]) }
                    }
                    if (cell.value && typeof cell.value == 'string') {
                        let curretValue = cell.value.toString()
                        cell.value = curretValue.replace(replace_regex, obj.dataFile[data_key]);
                    }
                });
            });
        }
    })

    let buffer_res = await workbook.xlsx.writeBuffer()
    zip.addFile(path.join(`${obj.folderName}`, obj.fileName + '.xlsx'), buffer_res);
}

async function sendFileToEmail(zip_path, email) {
    let mailConfig;
    
    mailConfig = {
        host: process.env.SMTP_HOST,
        port: process.env.SMTP_PORT, 
        secure: true,
        authMethod: 'LOGIN',
        auth: {
            user: process.env.SMTP_EMAIL,
            pass: process.env.SMTP_PASSWORD,
        }
    };
    
    logger.info("MAIL", "config: %s", mailConfig);

    let transporter = nodemailer.createTransport(mailConfig);

    let info = await transporter.sendMail({
        from: process.env.SMTP_EMAIL,
        to: email,
        subject: "BulkEditor App - Loaded Files",
        html: "<h3>Спасибо что воспользовались BulkEditor App!</h3>",
        attachments: [  
            {
                path: zip_path
            }
        ]
    });

    fs.unlinkSync(zip_path)
    logger.info("MAIL", "Message sent: %s", info.messageId);
}

module.exports = router;