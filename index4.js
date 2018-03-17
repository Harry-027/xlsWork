var xls = require('xlsx');
var path = require('path');

// creating a new workbook and worksheet and then writing a json data into workbook
// merging cells

function basicExcel() {
    var ws = xls.utils.json_to_sheet([
        { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
        { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 },
        {},{S: "No data avalable"}
    ], { header: ["S", "h", "e", "e_1", "t", "J", "S_1"], origin: "A1" });

    ws["!merges"] = [
        { s: { r: 4, c: 0 }, e: { r: 4, c: 6 } } /* A1:A2 */
    ]

    var wb = xls.utils.book_new();

    xls.utils.book_append_sheet(wb, ws, 'Sheet1');
    console.log(ws);
    var wopts = { bookType: 'xlsx', bookSST: false, type: 'binary' };
    xls.writeFile(wb, 'SheetJS.xlsx', wopts);
    var md = xls.readFile('SheetJS.xlsx', { cellStyles: true });
    console.log('md', md.Sheets.Sheet1);
}

basicExcel();
