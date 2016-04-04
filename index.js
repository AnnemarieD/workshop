var express = require('express');
var app = express();

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
workbook.xlsx.readFile('test.xlsx')
    .then(function() {
        // use workbook 
    });

app.get('/file', function(req, res) {
	res.send(workbook);
})
	
app.listen(2003, function () {
  console.log('Example app listening on port 2003!');
});