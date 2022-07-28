'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
    sourceFile: '/Users/barankaplan/Downloads/excelChart/template/GroupedBarChart.xlsx'
});

console.log(result);