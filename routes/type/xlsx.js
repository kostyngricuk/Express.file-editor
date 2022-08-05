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
                        console.log(`template_file - PATH "${file.filepath}" and MIMETYPE "${file.mimetype}"`,)
                        template_path = file.filepath
                        break;
                    case 'data_file':
                        console.log(`data_file - PATH "${file.filepath}" and MIMETYPE "${file.mimetype}"`,)
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