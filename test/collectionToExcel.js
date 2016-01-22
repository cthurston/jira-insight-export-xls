var expect = require('chai').expect;
var fs = require('fs');
var collectionToExcel = require('../collectionToExcel');

var collWithAttrs = require('./mock/insightCollection').withAttributes;

describe('Insight collection to XLS', function() {
	it('should create a valid .xls file with default options', function(done) {
		var options = {
			data: collWithAttrs
		};
		collectionToExcel(options, function(error, filePath, data) {
			expect(error).to.not.exist;
			expect(filePath).to.match(/^tmp\\export_.+xlsx$/);

			fs.access(filePath, fs.F_OK, function(err) {
				expect(err).to.not.exist;
				done();
			});

		});
	});

	it('should create a valid .xls file with set options', function(done) {
		var options = {
			data: collWithAttrs,
			path: './',
			name: 'testing'
		};
		collectionToExcel(options, function(error, filePath, data) {
			expect(error).to.not.exist;
			expect(filePath).to.match(/^testing.+xlsx$/);

			fs.access(filePath, fs.F_OK, function(err) {
				expect(err).to.not.exist;
				done();
			});

		});
	});

});