var xlsx = require('js-xlsx');
var path = require('path');

// creating a new workbook and worksheet and then writing a json data into workbook

function basicExcel() {
    var range = { s: { c: 1000, r: 1000 }, e: { c: 0, r: 0 } };
    var ws = {
        A1: { v: "ok", t: "s" }, B1: { v: "ok", t: "s" }, C1: { v: "ok", t: "s" }, D1: { v: "ok", t: "s" }, A2: { v: "ok", t: "s" }
    }
    ws['!ref'] = 'A1:D3';
    for (key in ws) {
        if (key != '!ref')
            ws[key].s = {
                font: {
                    bold: true
                },
                fill: {
                    fgColor: { rgb: "FFFFAA00" }
                }
            }
        console.log('ws', ws[key]);
    }
    var wb = new Workbook();

    ws_name = 'Sheet2'
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    xlsx.writeFile(wb, 'SheetJS.xlsx');
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

basicExcel();
