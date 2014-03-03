Excel.js
========

Native node.js Excel file parser. Only supports xlsx for now.

Install
=======
    npm install excel

Use
====
    var parseXlsx = require('excel');

    parseXlsx('Spreadsheet.xlsx', function(err, data) {
    	if(err) throw err;
        // data is an array of arrays of object
    });
    
If you have multiple sheets in your spreadsheet, 

    parseXlsx('Spreadsheet.xlsx', '2', function(err, data) {
    	if(err) throw err;
        // data is an array of arrays of object
    });

If you want to get multiple sheets from your spreadsheet,
set to argument start and end of your spreadsheet
    
    var start = 1;
    var end = 5;

    parseXlsx('Spreadsheet.xlsx', start, end, function(err, data) {
    	if(err) throw err;
        // data is an array of arrays of object
    
    });
    
MIT License.

*Author: Trevor Dixon <trevordixon@gmail.com>*

Contributors: 
- Jake Scott <scott.iroh@gmail.com>
- Fabian Tollenaar <fabian@startingpoint.nl> (Just a small contribution, really)
- amakhrov
