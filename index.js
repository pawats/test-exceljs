var express = require('express');
var Excel = require('exceljs');
var fs = require('fs');

var app = express();


app.get('/file', function (req, res) {
	var myColumns = ['Name', 'Country'];
	var myRows = [
					myColumns, 
					['Pat', 'Thailand']
				];

	var filename = 'output.xlsx';

	createExcelFile('MySheet', myRows, filename).then(function(){
		res.setHeader('Content-disposition', 'attachment; filename='+filename);
	    var stream = fs.createReadStream(__dirname + '/' + filename);
	    stream.pipe(res);
		// res.sendFile(filename, {root: './'});	
	})
});


app.get('/stream', function (req, res) {
	var myColumns = ['Name', 'Country'];
	var myRows = [
					myColumns, 
					['Pat', 'Thailand']
				];

	var filename = 'output.xlsx';

	//https://github.com/guyonroche/exceljs/issues/150
	res.setHeader('Content-disposition', 'attachment; filename='+filename);
	createExcelStream('MySheet', myRows, res);

});


app.listen(3000, function () {
  console.log('Example app listening on port 3000!');
});



function createExcelFile(sheetName, rows, filename){
	var workbook = new Excel.Workbook();
	workbook.created = new Date();
	workbook.modified = new Date();

	workbook.views = [
	  {
	    x: 0, y: 0, width: 10000, height: 20000, 
	    firstSheet: 0, activeTab: 1, visibility: 'visible'
	  }
	];

	var sheet = workbook.addWorksheet(sheetName, {properties:{tabColor:{argb:'FFC0000'}}});
	var worksheet = workbook.getWorksheet(sheetName);
	worksheet.addRows(rows);

	return workbook.xlsx.writeFile(filename)
};



function createExcelStream(sheetName, rows, stream){
	// construct a streaming XLSX workbook writer with styles and shared strings
	var options = {
	    // filename: './'+filename,
	    stream: stream,
	    useStyles: true,
	    useSharedStrings: true
	};

	var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
	workbook.created = new Date();
	workbook.modified = new Date();

	workbook.views = [
	  {
	    x: 0, y: 0, width: 10000, height: 20000, 
	    firstSheet: 0, activeTab: 1, visibility: 'visible'
	  }
	];

	var sheet = workbook.addWorksheet(sheetName, {properties:{tabColor:{argb:'FFC0000'}}});
	var worksheet = workbook.getWorksheet(sheetName);

	rows.forEach(function(r){
		worksheet.addRow(r).commit();		
	})

	worksheet.commit();
	workbook.commit();
	return workbook;


};