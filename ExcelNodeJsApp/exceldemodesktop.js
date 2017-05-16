var Excel = require('@microsoft/office-js/excel');
var exceldemolib = require('./exceldemolib.js')
var OfficeExtension = require('@microsoft/office-js/office.runtime');

OfficeExtension.Utility._logEnabled = true;


OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = {url: "http://localhost:8052"};
var p = new OfficeExtension.Promise(function(resolve, reject){
    resolve(null);
});
    p.then(function(){
        return exceldemolib.clearWorkbook();
    })
    .then(function(){
        return exceldemolib.dataPopulateSetup()
    })
    .then(function(){
        console.log("invoking dataPopulateRun");
        return exceldemolib.dataPopulateRun();
    })
    .then(function(){
        console.log("invoking getChartImage");
        return exceldemolib.getChartImage();
    })
    .then(function(imageData){
        console.log("---ImageData start---");
        console.log(imageData);
        console.log("---ImageData end---");
    })
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });

