var Excel = require('exceljs');

var createExcelFile = function(sheetName, rows, filename){
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



var createExcelStream = function(sheetName, rows, stream){
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

module.exports = {
	createExcelFile: createExcelFile, 
	createExcelStream: createExcelStream
}