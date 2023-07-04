var conf = require("config");

var excelfunc = require("./excelfunc.js");


if (conf.has("ExcelFilesCatalog")) {
    console.log(conf.get("ExcelFilesCatalog"));
    excelfunc.CreateTestExcelFiles(conf.get("ExcelFilesCatalog"))
    
}else {
    console.log("Not find fucking catalog!");
    return;
}


