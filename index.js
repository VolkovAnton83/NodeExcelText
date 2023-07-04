var conf = require("config");
if (conf.has("ExcelFilesCatalog")) {
    console.log(conf.get("ExcelFilesCatalog"));
}else {
    console.log("Not find fucking catalog!");
}
