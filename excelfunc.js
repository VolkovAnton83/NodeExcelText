const fs = require('fs');
var writeexcel = require('excel4node');
var readexcel = require('node-xlsx');
const { exit } = require('process');
//ExcelJS for read? 2k packets. 
//node-xlsx 46 packets only

module.exports = { CreateTestExcelFiles }

function CreateExcelList(wrkbook,wrksheetname) {
    var worksheet = wrkbook.addWorksheet(wrksheetname);
    worksheet.cell(1,1).string(Date.now().toString());
}

function CreateTestExcelFiles(FolderPath,TemplateName) {
    
    const workSheetsFromFile = readexcel.parse(FolderPath+"\\"+TemplateName);
    console.log(typeof workSheetsFromFile);
    console.log(JSON.stringify(workSheetsFromFile, null, 6));
    return;
    
    var workbook = new writeexcel.Workbook();
    const t0 = performance.now();
    
    var curDate = new Date('01/01/2024 00:00:01');
    while (curDate < new Date('12/31/2024 00:00:01')) {
        curDate.setDate(curDate.getDate() + 1);
        if ((curDate.getDay() == 1) || (curDate.getDay() == 3) ) {
            CreateExcelList(workbook,curDate.toLocaleDateString('ru-RU'));
        }
        //console.log('' + curDate.toLocaleDateString('ru-RU') +' - '+ curDate.getDay());
    }
    console.log(`Call to doSomething took ${(performance.now() - t0)} milliseconds.`);


    // var worksheet = workbook.addWorksheet('Sheet 1');
    // var worksheet2 = workbook.addWorksheet('Sheet 2');
   
    // var style = workbook.createStyle({
    // font: {
    //     color: '#FF0800',
    //     size: 12
    // },
    // numberFormat: '$#,##0.00; ($#,##0.00); -'
    // });

    // worksheet.cell(1,1).number(100).style(style);

    // worksheet.cell(1,2).number(200).style(style);

    // worksheet.cell(1,3).formula('A1 + B1').style(style);

    // worksheet.cell(2,1).string('string').style(style);

    // worksheet.cell(3,1).bool(true).style(style).style({font: {size: 14}});

    //const xmas95 = new Date("1995-12-25T23:15:30");
    
    
    //console.log(seconds); // 30

    // worksheet.cell(1,1).string(Date.now().toString());

    workbook.write(FolderPath+'\\Журнал поступлений на '+curDate.getFullYear()+' год.xlsx');

}