"use strict";
var fs = require('fs');
var process = require('process');
var workingDirectory = process.cwd().slice(2);
var XLSX = require('xlsx');
var workbook = XLSX.readFile(process.argv[2]);
var sheets = workbook.Sheets;
var htmlFile = '';
var rowNumber;
var htmlArray;

// Check to make sure user provides argument for command line
if (typeof process.argv[2] === 'undefined') {
	console.log('\n' + 'Error:' + '\n' + 'You must enter the excel file you wish to build tables from as an argument' + '\n' + 'i.e., node toTable.js resolutions.xlsx');
	return;
} else {
	// Check that the file is the correct type
	if (process.argv[2].slice(-4) !== 'xlsx') {
		console.log('\n' + 'This program will only convert xlsx files' + '\n' + 'Please enter correct file type');
		return;
	} else {
		// Create the HTML file name to write the table to
		var fileName = process.argv[2];
		var newFileName = fileName.slice(0, -4) + 'html';
	}
}

htmlFile += '<html>' + '\n' + '<body>' +'\n';

function getPosition(string, subString, index) {
   return string.split(subString, index).join(subString).length;
}
// Iterate through each worksheet in the workbook
for (var sheet in sheets) {
	
	// Start building a new table if the worksheet has entries
	if (typeof sheet !== 'undefined') {
		htmlFile += '<table>' + '\n';		
		// Iterate over each cell value on the sheet
		for (var cell in sheets[sheet]) {			
							
			// Protect against undefined values
			if (typeof sheets[sheet][cell].w !== 'undefined') {
				if (cell.slice(0, 1) === 'A') htmlFile += '<tr>';
				if (cell.slice(0, 1) === 'A' || cell.slice(0, 1) === 'B' || cell.slice(0, 1) === 'C') {
					htmlFile += '<td>' + sheets[sheet][cell].w.replace('&', '&amp;').replace('<', '&lt;') + '</td>';
				}
				if (cell.slice(0, 1) === 'C') htmlFile += '</tr>' + '\n';
			}
		}
		// Close the table
		htmlFile += '</table>' + '\n';
	}
	/*console.log(sheets[sheet]['!merges']);
	sheets[sheet]['!merges'].forEach(function(merge, index) {
		//console.log(merge);
		rowNumber = (getPosition(htmlFile, '<th>', (merge.s.r+1)) + 3);
		console.log(rowNumber);
		htmlArray = htmlFile.split('');
		htmlArray.splice(rowNumber, 0, ' colspan="3"');
		htmlFile = htmlArray.join('');
	});*/
}
// Close the file
htmlFile += '</body>' + '\n' + '</html>';

// Write htmlFile variable to the disk with newFileName as the name
fs.writeFile(newFileName, htmlFile, (err) => {
	if (err) throw err;
	console.log('\n' +'Your tables have been created in', newFileName);
});
