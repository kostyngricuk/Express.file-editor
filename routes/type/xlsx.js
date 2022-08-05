const express = require('express');
const router = express.Router();

const multer = require('multer')
const storage = multer.memoryStorage()
const upload = multer({ storage: storage }).single("upload_file")
const XLSX = require('xlsx')

/* GET xlsx */
router.route('/')
    .post((req, res) => {
        upload(req, res, (err) => {
            if (err) {
                res.status(400).send("Something went wrong!");
            }

            workbook = XLSX.read(req.file.buffer, {
                type: 'buffer',
            })

            let worksheet = workbook.Sheets[workbook.SheetNames[0]]

            let worksheetJson = XLSX.utils.sheet_to_json(worksheet)

            let worksheetString = JSON.stringify(worksheetJson);

            // UPDATE DATA
            let index = '001'
            let day = '2'
            let month = 'августа'
            let year = '2022'
            let number_plate = 'AS-2312'
            let number_garage = '4Д'
            worksheetString = worksheetString.replaceAll('#index#', index)
            worksheetString = worksheetString.replaceAll('#day#', day)
            worksheetString = worksheetString.replaceAll('#month#', month)
            worksheetString = worksheetString.replaceAll('#year#', year)
            worksheetString = worksheetString.replaceAll('#number_plate#', number_plate)
            worksheetString = worksheetString.replaceAll('#number_garage#', number_garage)
            // ---

            let resJson = JSON.parse(worksheetString)

            res.send(resJson);


            // const buf = XLSX.write(workbook, { type:"buffer", bookType:"xlsx" });
            // res.statusCode = 200;
            // res.setHeader('Content-Disposition', 'attachment; filename="result.xlsx"');
            // res.setHeader('Content-Type', 'application/vnd.ms-excel');
            // res.end(buf);
        })
    })

module.exports = router;