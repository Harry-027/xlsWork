var Excel = require('exceljs');
var workbook = new Excel.Workbook();

// workbook.creator = 'Me';
// workbook.lastModifiedBy = 'Her';
// workbook.created = new Date(1985, 8, 30);
// workbook.modified = new Date();
// workbook.lastPrinted = new Date(2016, 9, 27);

// workbook.views = [
//     {
//         x: 0, y: 0, width: 10000, height: 20000,
//         firstSheet: 0, activeTab: 1, visibility: 'visible'
//     }
// ]

var worksheet = workbook.addWorksheet('Active', {
    pageSetup: { paperSize: 9, orientation: 'landscape' }
});

// add image to workbook by filename
var imageId1 = workbook.addImage({
    filename: './s.png',
    extension: 'png'
});

worksheet.addBackgroundImage(imageId1, {
    tl: { col: 0.1125, row: 0.4 },
    br: { col: 2.101046875, row: 3.4 },
    editAs: 'absolute'
});

worksheet.autoFilter = 'A1:C1';
worksheet.columns = [
    {
        header: 'Id', key: 'id', width: 10, style: {
            font: {
                name: 'Arial Black', size: 16,
                underline: true,
                bold: true,
                color: { argb: 'FF00FF00' }
            }
        }
    },
    {
        header: 'Name', key: 'name', width: 32, style: {
            font: {
                name: 'Arial Black', size: 16,
                underline: true,
                bold: true,
                color: { argb: 'FF00FF00' }
            },
            alignment: { vertical: 'middle', horizontal: 'center', textRotation: -45 },
            border: {
                top: { style: 'double', color: { argb: 'FF00FF00' } },
                left: { style: 'double', color: { argb: 'FF00FF00' } },
                bottom: { style: 'double', color: { argb: 'FF00FF00' } },
                right: { style: 'double', color: { argb: 'FF00FF00' } }
            },
            fill: {
                type: 'pattern',
                pattern: 'darkVertical',
                fgColor: { argb: 'FFFF0000' }
            }
        }
    },
    {
        header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1, style: {
            font: {
                name: 'Arial Black', size: 16,
                underline: true,
                bold: true,
                color: { argb: 'FF00FF00' }
            },
            fill: {
                type: 'gradient',
                gradient: 'angle',
                degree: 0,
                stops: [
                    { position: 0, color: { argb: 'FF0000FF' } },
                    { position: 0.5, color: { argb: 'FFFFFFFF' } },
                    { position: 1, color: { argb: 'FF0000FF' } }
                ]
            }
        }
    }
];

var row = worksheet.getRow(4);
worksheet.mergeCells('A3:B8');
var rows = [
    [5, 'Bob', new Date()], // row by array
    { id: 6, name: 'Barbara', dob: new Date() }
];
worksheet.addRows(rows);

workbook.xlsx.writeFile('Sheet3.xlsx')
    .then(function () { });