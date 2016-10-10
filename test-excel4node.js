var xl = require ('excel4node');


var generate = function(filename, data, mergeCells, res){
	// Create a new instance of a Workbook class
	var wb = new xl.Workbook();

	var options = {
	    margins: {
	        left: 1.5,
	        right: 1.5
	    }
	};
	// Add Worksheets to the workbook
	var ws = wb.addWorksheet('Sheet 1', options);
	// var ws2 = wb.addWorksheet('Sheet 2');

	// Create a reusable style
	var style = wb.createStyle({
	    font: {
	        color: '#000000',
	        size: 12
	    },
	    // numberFormat: '$#,##0.00; ($#,##0.00); -'
	});

	// // Set value of cell A1 to 100 as a number type styled with paramaters of style
	// ws.cell(1,1).number(100).style(style);

	// // Set value of cell B1 to 300 as a number type styled with paramaters of style
	// ws.cell(1,2).number(200).style(style);

	// // Set value of cell C1 to a formula styled with paramaters of style
	// ws.cell(1,3).formula('A1 + B1').style(style);

	// // Set value of cell A2 to 'string' styled with paramaters of style
	// ws.cell(2,1).string('string').style(style);

	// // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
	// ws.cell(3,1).bool(true).style(style).style({font: {size: 14}});

	if(mergeCells){
		for(var c = 0; c < mergeCells.length; c++){
			var mc = mergeCells[c];
			ws.cell(mc.startRow, mc.startColumn, mc.endRow, mc.endColumn, true);
		}		
	}


	//Go through each row in data array
	for(var row = 0; row < data.length; row++){
		for(var col = 0; col < data[row].length; col++){
			var currentCell = ws.cell(row + 1, col + 1);
			var currentCellData = data[row][col];
			switch(typeof currentCellData.value){
				case "string":
					currentCell.string(currentCellData.value)					
					break;

				case "number":
					currentCell.number(currentCellData.value)					
					break;				
			}

			if(currentCellData.style){
				currentCell.style(currentCellData.style)
			}
			else{
				currentCell.style(style)
			}
		}
	}

	var myStyle = wb.createStyle({
	    font: {
	        bold: true,
	        color: '00FF00'
	    }
	});


	ws.column(3).setWidth(25);



	var filename = (filename) ? (filename + '.xlsx') : ('output.xlsx');
	wb.write(filename, res);

}



module.exports = {
	generate: generate
}