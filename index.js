var express = require('express');
var app = express();

if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
var data = JSON.stringify(to_json(workbook), 2, 2);

app.get('/users', function (req, res) {
  res.send(data.map({desired_value}));
});

var first_sheet_name = workbook.SheetNames[0];
var address_of_cell = 'A2';
 
/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];
 
/* Find desired cell */
var desired_cell = worksheet[address_of_cell];
 
/* Get the value */
var desired_value = desired_cell.v;



app.get('/userlist', function(req, res) {
	//print user list
	res.send(column);
})

app.get('/file', function(req, res) {
	//prints complete file
	res.send(data);
})
	
app.listen(2003, function () {
  console.log('Example app listening on port 2003!');
});

function to_json(wb) {
	var result = {};
	wb.SheetNames.forEach(function(sheetName) {
		var roa = XLSX.utils.sheet_to_row_object_array(wb.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
};