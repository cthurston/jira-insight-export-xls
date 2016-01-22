var _ = require('lodash');
var path = require('path');
var humanize = require('underscore.string').humanize;
var striptags = require('striptags');
var excelBuilder = require('msexcel-builder');
var unwrapCollection = require('jira-insight-api-wrapper').unwrapCollection;

module.exports = collectionToExcel;

/**
 * Will call next with the error, full path to the file, and the unwrapped data.
 * @param {Object} options
 * @param {Object[]} options.data - The Jira Insight Collection of objects.
 * @param {string} [options.name=export_(Date.now())] - The name of the xlsx file.
 * @param {string} [options.path=tmp] - The folder in which to save the file.
 * @param next
 * @returns undefined
 */
function collectionToExcel(options, next) {
	var data = unwrapCollection(options.data);
	var name = options.name || 'export_' + Date.now();
	var folder = options.path || './tmp';
	var filename = name + '.xlsx';

	var rows;
	var cols;
	var headings;
	var workbook;
	var sheet1;

	if (!data.objectEntries.length) {
		return next(new Error('There are no objects to export'));
	}

	rows = data.objectEntries.length + 2;
	headings = Object.keys(data.objectEntries[0]);
	cols = headings.length;

	workbook = excelBuilder.createWorkbook(folder, filename);
	sheet1 = workbook.createSheet('sheet1', cols, rows);

	//Fill in headers:
	_.each(headings, function(th, index) {
		sheet1.set(index + 1, 1, humanize(th));
	});

	//Fill in the data:
	_.each(data.objectEntries, function(obj, r) {
		_.each(headings, function(key, c) {
			sheet1.set(c + 1, r + 2, parseValue(obj[key]));
		});
	});

	// Save it
	workbook.save(function(error, jszip) {
		if (error) {
			next(error);
		} else {
			next(error, path.join(folder,filename), data);
		}
	});
}

function parseValue(obj) {
	if (_.isArray(obj)) {
		return (_.map(obj, parseValue)).toString();
	} else if (_.isNull(obj)) {
		return '';
	} else if (_.isObject(obj) && obj.name) {
		return obj.name;
	} else if (_.isObject(obj)) {
		return JSON.stringify(obj);
	} else if (_.isString(obj)) {
		return striptags(obj);
	} else {
		return obj;
	}
}