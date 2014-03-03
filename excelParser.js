var Promise = require('node-promise'),
	defer = Promise.defer,
	when = Promise.when,
	all = Promise.all,
	_ = require('underscore');

function extractFiles(path, start, end) {
    var unzip = require('unzip');
    var deferred = defer();

    var files = {
        strings: {
            deferred: defer()
        },
        book: {
            deferred: defer()
        },
        'xl/sharedStrings.xml': 'strings',
        'xl/workbook.xml': 'book'
    };
    for (var i = start; i <= (end || start); i++) {
        files['sheet' + i] = {
            deferred: defer()
        };
        files['xl/worksheets/sheet' + i + '.xml'] = 'sheet' + i;
    }

    var srcStream = path instanceof require('stream') ?
		path :
		require('fs').createReadStream(path);

    srcStream
    .pipe(unzip.Parse())
    .on('error', function(err) {
        deferred.reject(err);
    })
    .on('entry', function(entry) {
        if (files[entry.path]) {
            var contents = '';
            entry.on('data', function(data) {
                contents += data.toString();
            }).on('end', function() {
                files[files[entry.path]].contents = contents;
                files[files[entry.path]].deferred.resolve();
            });
        }
    });

	when(all(_.pluck(files, 'deferred')), function() {
		deferred.resolve(files);
	});

	return deferred.promise;
}

function extractData(files) {
	var libxmljs = require('libxmljs'),
		book = libxmljs.parseXml(files.book.contents),
		strings = libxmljs.parseXml(files.strings.contents),
		ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'},
        sheets = {};
 
    _.each(files, function(sheet, key) {
        if (key.match(/^sheet[0-9]*$/)) {
            sheets[key] = sheet;
        }
    });

	var colToInt = function(col) {
		var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
		var col = col.trim().split('');
		
		var n = 0;

		for (var i = 0; i < col.length; i++) {
			n *= 26;
			n += letters.indexOf(col[i]);
		}

		return n;
	};

	var Cell = function(cell) {
		cell = cell.split(/([0-9]+)/);
		this.row = parseInt(cell[1]);
		this.column = colToInt(cell[0]);
	};


    //book data
    var b = book.find('//a:sheets//a:sheet', ns);
    var names = _.map(b, function(tag) {
        return tag.attr('name').value();
    });

    //xml data
    var xmls = {};
    _.each(names, function(name, i) {
        if (_.has(sheets, 'sheet' + (i+1))) {
            xmls[name] = sheets['sheet' + (i+1)];
        }
    });

    var result = [];
    _.each(xmls, function(xml, name) {
		var sheet = libxmljs.parseXml(xml.contents);
        var d = sheet.get('//a:dimension/@ref', ns).value().split(':');
		var data = [];

        d = _.map(d, function(v) { return new Cell(v); });

        var cols = d[1].column - d[0].column + 1,
        rows = d[1].row - d[0].row + 1;

        _(rows).times(function() {
            var _row = [];
            _(cols).times(function() { _row.push(''); });
            data.push(_row);
        });

        var cells = sheet.find('//a:sheetData//a:row//a:c', ns),
        na = {
            value: function() { return ''; },
            text:  function() { return ''; }
        };
        _.each(cells, function(_cell) {
            var r = _cell.attr('r').value(),
            type = (_cell.attr('t') || na).value(),
            value = ( _cell.get('a:v', ns) || na ).text(),
            cell = new Cell(r);
            if (type == 's') {
                value = strings.get('//a:si[' + (parseInt(value) + 1) + ']//a:t', ns).text();
            }
            data[cell.row - d[0].row][cell.column - d[0].column] = value;
        });
        result.push({
            name: name,
            data: data
        });
    });
	return result;
}

module.exports = function parseXlsx() {
    var path, sheet, cb, start, end;
    if(arguments.length == 2) {
        path = arguments[0];
        sheet = 1;
        cb = arguments[1];
    }
    if(arguments.length == 3) {
        path = arguments[0];
        sheet = arguments[1];
        cb = arguments[2];
    }
    if (arguments.length == 4) {
        path = arguments[0];
        start = arguments[1];
        end = arguments[2];
        cb = arguments[3];
    }
    extractFiles(path, sheet || start, end || null).then(function(files) {
        cb(null, extractData(files));
    },
    function(err) {
        cb(err);
    });
};
