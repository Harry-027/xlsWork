var Excel = require('exceljs');
const data = require('./data.js')

var workbook = new Excel.Workbook();
var arr = [];
var rows = [];

workbook.views = [
    {
        x: 0, y: 0, width: 10000, height: 20000,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
]

var activeWorksheet = workbook.addWorksheet('Active', {
    pageSetup: {
        paperSize: 9, orientation: 'portrait',
        horizontalCentered: true, verticalCentered: true
    }, properties: { tabColor: { argb: 'FF00FF00' } }
})

var archiveWorksheet = workbook.addWorksheet('Archive', {
    pageSetup: {
        paperSize: 9, orientation: 'portrait',
        horizontalCentered: true, verticalCentered: true
    }, properties: { tabColor: { argb: 'FF00FF00' } }
})

const defaultWidth = 10;
const defaultStyle = {
    font: {
        name: 'Arial Black', size: 12,
        bold: true
    }
}
const defaultAlignment = { vertical: 'middle', horizontal: 'center' }
const stageList = [{ key: 'stage', value: 'Stage' }, { key: 'name', value: 'Name' }, { key: 'askingprice', value: 'Asking Price' }];

for (let i = 0; i < 10; i++) {
    arr.push({});
}
addConfig().then(() => {
    addData().then(() => {
        activeWorksheet.addRows(rows);
        archiveWorksheet.addRows(rows);
        activeWorksheet.getRow(14).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', italic: true };
        activeWorksheet.getRow(14).commit();
        workbook.xlsx.writeFile('Sheet3.xlsx')
            .then(function () { });
    })
})

activeWorksheet.spliceColumns(1, 0, ...arr);
activeWorksheet.spliceRows(1, 0, ...arr);

archiveWorksheet.spliceColumns(1, 0, ...arr);
archiveWorksheet.spliceRows(1, 0, ...arr);




function addConfig() {
    var columns = [];
    return new Promise((res, rej) => {
        stageList.forEach((data) => {
            columns.push({
                header: data.value,
                key: data.key,
                width: defaultWidth,
                style: defaultStyle,
                alignment: defaultAlignment
            });
        });
        activeWorksheet.columns = columns;
        archiveWorksheet.columns = columns;
        res();
    });
}

function addData() {
    return new Promise((res, rej) => {
        for (var key in data.data) {
            data.data[key].forEach(function (element) {
                rows.push({
                    stage: element.stage,
                    name: element.name,
                    askingprice: element.askingprice
                });
            });
        }
        res();
    });
}