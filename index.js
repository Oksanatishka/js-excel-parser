var XLSX = require('xlsx');

function excel_to_json(excelPath) {
    var json_sheet = [];
    var workbook = XLSX.readFile(excelPath);
    var sheets = workbook.SheetNames;
    for (var i = 0; i < sheets.length; i++) {
        var ws;
        try {
            ws = workbook.Sheets[sheets[i]];
            json_sheet = XLSX.utils.sheet_to_json(ws);
        } catch (e) {
            console.error('error parsing: ' + e);
        }
    }
    return json_sheet;
}
module.exports = excel_to_json;
