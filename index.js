var express = require('express');
var fs = require('fs');
var bodyParser = require('body-parser');
var testExceljs = require('./test-exceljs.js');
var testExcel4Node = require('./test-excel4node.js');

var app = express();

app.use(bodyParser.json({
	limit: '50mb'
}));       // to support JSON-encoded bodies

app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true,
  limit: '50mb'
})); 

app.use(express.static('view'));


app.post('/excel4node/generate', function(req, res){
	if(req.body.data.length > 0){
		var data = JSON.parse(req.body.data);
		var filename = req.body.filename ? req.body.filename : undefined;
		var mergeCells = req.body.mergecells ? JSON.parse(req.body.mergecells) : undefined;

		testExcel4Node.generate(filename, data, mergeCells, res);			
	}else{
		res.end('no input data!')
	}
	
})


app.post('/exceljs/stream', function (req, res) {
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
		testExceljs.createExcelStream('MySheet', data, res);

	}
	
	else
	{
		res.end('Please enter data')
	}

});


app.post('/exceljs/file', function (req, res) {
	var myColumns = ['Name', 'Country'];
	var myRows = [
					myColumns, 
					['Pat', 'Thailand']
				];

	var filename = 'output.xlsx';

	testExceljs.createExcelFile('MySheet', myRows, filename).then(function(){
		res.setHeader('Content-disposition', 'attachment; filename='+filename);
	    var stream = fs.createReadStream(__dirname + '/' + filename);
	    stream.pipe(res);
		// res.sendFile(filename, {root: './'});	
	})
});



app.use(function(err, req, res, next) {
  console.error(err.stack);
  res.status(500).send(String(err));
});

app.listen(3000, function () {
  console.log('Example app listening on port 3000!');
});


