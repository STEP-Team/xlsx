# xlsx

Parser and writer for various spreadsheet formats.

see <https://github.com/SheetJS/js-xlsx/blob/master/README.md>

This node module is derived from <https://www.npmjs.com/package/xlsx> v.0.8.0

This module is just an experiment and we do not plan any specific maintenance.

The code adds an extra attribute 'XML' to Workbook so Workbook specialized structure (e.g. definedNames) is accessible like this:

```
'use strict';

var express = require('express');
var XLSX = require('xlsx');
var xpath = require('xpath.js');
var dom = require('xmldom').DOMParser;

var excelObj = XLSX.readFile("./myExcelFile.xlsx");
var worksheet = excelObj.Sheets[excelObj.SheetNames[0]];
var xml = "" + excelObj.Workbook.XML;
var doc = new dom().parseFromString(xml);

// e.g. Fetching the range for a name COUNTRY defined in the first worksheet; using 'local-name()' to work around namespaces
var range = xpath(doc, "//*[@name='COUNTRY' and local-name(.)='definedName']")[0].firstChild.toString();
console.log(range);
```
