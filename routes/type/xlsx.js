var express = require('express');
var router = express.Router();

/* GET xlsx */
router.route('/')
    .get(function(req, res) {
        res.end('GET');
    })
    .post(function(req, res) {
        res.end('POST')
    })
    .put(function(req, res) {
        res.end('PUT')
    })
    .delete(function(req, res) {
        res.end('DELETE')
    })

module.exports = router;