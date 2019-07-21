var express = require('express');
var app = express();
const excel = require('exceljs');

const jsonCustomers = 
		
			[ { id: 1, address: 'Jack Smith', age: 23, name: 'Massachusetts' },
			{ id: 2, address: 'Adam Johnson', age: 27, name: 'New York' },
			{ id: 3, address: 'Katherin Carterg', age: 26, name: 'Washington DC' },
			{ id: 4, address: 'Jack London', age: 33, name: 'Nevada' },
			{ id: 5, address: 'Jason Bourne', age: 36, name: 'California' },
			{ id: 1, address: 'Jack Smith', age: 23, name: 'Massachusetts' },
			{ id: 2, address: 'Adam Johnson', age: 27, name: 'New York' },
			{ id: 3, address: 'Katherin Carterg', age: 26, name: 'Washington DC' },
			{ id: 4, address: 'Jack London', age: 33, name: 'Nevada' },
			{ id: 5, address: 'Jason Bourne', age: 36, name: 'California' },
			{ id: 1, address: 'Jack Smith', age: 23, name: 'Massachusetts' },
			{ id: 2, address: 'Adam Johnson', age: 27, name: 'New York' },
			{ id: 3, address: 'Katherin Carterg', age: 26, name: 'Washington DC' },
			{ id: 4, address: 'Jack London', age: 33, name: 'Nevada' },
			{ id: 5, address: 'Jason Bourne', age: 36, name: 'California' } ]
	
		
		let workbook = new excel.Workbook(); //creating workbook
		let worksheet = workbook.addWorksheet('Customers'); //creating worksheet
	 
		worksheet.pageSetup.margins = {
			left: 5, right: 5,
			top: 5, bottom: 5,
			header: 3, footer: 3
		  };

		//  WorkSheet Header
		worksheet.columns = [
			{ header: 'Id', key: 'id', width: 10},
			{ header: 'Name', key: 'name', width: 30 },
			{ header: 'Address', key: 'address', width: 30},
			{ header: 'Age', key: 'age', width: 10, outlineLevel: 1}
		];
	 
		// Add Array Rows
		worksheet.addRows(jsonCustomers);
	 


		app.get('/download',function(req,res)
		{
		var tempfile = require('tempfile');
		var tempFilePath = tempfile('.xlsx');
		console.log("tempFilePath : ", tempFilePath);
		workbook.xlsx.writeFile(tempFilePath).then(function() {
		res.sendFile(tempFilePath, function(err)
		{
        console.log('file Successfully get downloaded');
    });
});
		});

		var server = app.listen(8080, function () 
		{
			var host = server.address().address
			var port = server.address().port
			console.log("Example app listening at http://%s:%s", host, port)
		 })
		
		
		// -> Check 'customer.csv' file in root project folder
