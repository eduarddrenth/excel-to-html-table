# excel-to-html-table

This Fork's changes:

* only process first 3 columns in first sheet
* just plain table cells, no headings
* wrap in html/body
* entities only or & and <
* deal with missing cels in third column
* search function (help text in frisian)

## usage

* clone this repo
* cd into the directory
* install npm
* install nodejs
* npm install xlsx
* node index.js "path to your excel"


This program will convert each worksheet in and excel workbook into an HTML5 table and replace common symbols with their unicode 
equivalent(&amp;, &mdash;, &ndash)

The first row on each sheer will be formated as a table header and the first coloumn will be a th header cell for each row.

To use the program:

1. Move the excel workbook to the folder.
2. Open git bash and cd into the excel-to-html-table directory.
3. Run node index.js workbook.xlsx where the workbook.xlsx is the name of the excel workbook you would like to create the table from.
