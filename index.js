var xls = require('xlsx');
var path = require('path');

// creating a new workbook and worksheet and then writing a json data into workbook

function basicExcel() {
    var ws = xls.utils.json_to_sheet([
        { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
        { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 }
    ], { header: ["S", "h", "e", "e_1", "t", "J", "S_1"], origin: "A6" });

var wb = xls.utils.book_new();

// set the workbook properties

if (!wb.Props) wb.Props = {};
wb.Props.Title = "My new Book";
wb.Props.Author = "Myself";
wb.Props.CreatedDate = new Date();

// link to certain site

ws['A6'].l = { Target: "http://sheetjs.com", Tooltip: "Find us @ SheetJS.com!" };

xls.utils.book_append_sheet(wb, ws, 'Sheet1');
console.log(ws);
xls.writeFile(wb, 'SheetJS.xlsx');
}

basicExcel();
