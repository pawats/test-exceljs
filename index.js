var express = require('express');
var Excel = require('exceljs');
var fs = require('fs');
var bodyParser = require('body-parser');

var app = express();

app.use(bodyParser.json({
	limit: '50mb'
}));       // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true,
  limit: '50mb'
})); 
app.use(express.static('view'));


app.post('/file', function (req, res) {
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


app.post('/stream', function (req, res) {
	var data;
	var filename = "output.xlsx";

	if(req.body.data.length > 0)
	{
		data = JSON.parse(req.body.data);

		if(req.body.filename){
			filename = req.body.filename + '.xlsx';	
		}

		//https://github.com/guyonroche/exceljs/issues/150
		res.setHeader('Content-disposition', 'attachment; filename=' + filename);
		createExcelStream('MySheet', data, res);

	}
	
	else
	{
		res.end('Please enter data')
	}

});

app.use(function(err, req, res, next) {
  console.error(err.stack);
  res.status(500).send('Something broke!');
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