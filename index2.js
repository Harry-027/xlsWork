var xlsx = require('js-xlsx');
var xls = require('xlsx');
var path = require('path');

// creating a new workbook and worksheet and then writing a json data into workbook

function basicExcel() {
    console.log('xlsss', xlsx);
    // var ws = xls.utils.json_to_sheet([
    //     { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
    //     { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 }
    // ], { header: ["S", "h", "e", "e_1", "t", "J", "S_1"], origin: "A6" });
    var range = {s: {c:1000, r:1000}, e: {c:0, r:0 }};
    var ws = {
        A1: { v: "ok", t: "s" }, B1: { v: "ok", t: "s" }, C1: { v: "ok", t: "s" }, D1: { v: "ok", t: "s" }, A2: { v: "ok", t: "s" }
    }
    ws['!ref'] = 'A1:D3';
    var wb = new Workbook();

    // set the workbook properties

    if (!wb.Props) wb.Props = {};
    wb.Props.Title = "My new Book";
    wb.Props.Author = "Myself";
    wb.Props.CreatedDate = new Date();

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
    ws_name = 'Sheet2'
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    xlsx.writeFile(wb, 'SheetJS.xlsx');
}


function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

basicExcel();
