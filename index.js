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

htmlFile += '<!DOCTYPE html>' +'\n';
htmlFile += '<html>' + '\n' + '<body onkeypress="if (event.keyCode==13) findString()">' +'\n';

function getPosition(string, subString, index) {
   return string.split(subString, index).join(subString).length;
}
// Iterate through each worksheet in the workbook
for (var sheet in sheets) {
	
	// Start building a new table if the worksheet has entries
	if (typeof sheet !== 'undefined') {
		htmlFile += '<form id="f1" name="f1" action="javascript:void()" style="top: 10px;">';
		htmlFile += '<input id="t1" name="t1" value="" placeholder="syktekst en &lt;enter> (&lt;enter> werhelber)" size="40" type="text"> Sykje mei dit fjild wurket net in edge';
		htmlFile += '</form>';
		htmlFile += '<table style="margin-top: 50px; position: absolute">' + '\n';		
		// Iterate over each cell value on the sheet
		var closed = true;
		var row = 0;
		for (var cell in sheets[sheet]) {
			if (cell.slice(0, 1) === 'A') {
				if (!closed) htmlFile += '</tr>' + '\n';
				closed=false;
				htmlFile += '<tr' + ((++row % 2 == 0) ? '' : ' style="background-color: #eeeeee"') + '>';
			}
			// Protect against undefined values
			if (typeof sheets[sheet][cell].w !== 'undefined') {
				if (cell.slice(0, 1) === 'A' || cell.slice(0, 1) === 'B' || cell.slice(0, 1) === 'C') {
					htmlFile += '<td>' + sheets[sheet][cell].w.replace('&', '&amp;').replace('<', '&lt;') + '</td>';
				}
			}
			if (cell.slice(0, 1) === 'C') {htmlFile += '</tr>' + '\n'; closed=true;}
		}
		if (!closed) htmlFile += '</tr>' + '\n';
		// Close the table
		htmlFile += '</table>' + '\n';
	}
	break;
}
// Close the file
htmlFile += '<script type="text/javascript"><!--\n';
htmlFile += 'var t1=document.getElementById(\'t1\');\n';
htmlFile += 'function findString () {\n';
htmlFile += '    if (t1.value==null||t1.value==\'\') return;\n';
htmlFile += '        if (window.find) {\n';
htmlFile += '        if (!window.find(t1.value,false,null,true)) {\n';
htmlFile += '            t1.focus();\n';
htmlFile += '            document.body.scrollTop = document.documentElement.scrollTop = 0;\n';
htmlFile += '        }\n';
htmlFile += '    } else if (document.selection && document.body.createTextRange) {\n';
htmlFile += '        var sel = document.selection;\n';
htmlFile += '        var textRange;\n';
htmlFile += '        if (sel.type == "Text") {\n';
htmlFile += '            textRange = sel.createRange();\n';
htmlFile += '            textRange.collapse(false);\n';
htmlFile += '        } else {\n';
htmlFile += '            textRange = document.body.createTextRange();\n';
htmlFile += '        }\n';
htmlFile += '        if (textRange.findText(t1.value)) {\n';
htmlFile += '            textRange.select();\n';
htmlFile += '        } else {\n';
htmlFile += '            t1.focus();\n';
htmlFile += '            document.body.scrollTop = document.documentElement.scrollTop = 0;\n';
htmlFile += '        }\n';
htmlFile += '    }\n';
htmlFile += '}\n';
htmlFile += '--></script>\n';
htmlFile += '</body>' + '\n' + '</html>';

// Write htmlFile variable to the disk with newFileName as the name
fs.writeFile(newFileName, htmlFile, (err) => {
	if (err) throw err;
	console.log('\n' +'Your tables have been created in', newFileName);
});
