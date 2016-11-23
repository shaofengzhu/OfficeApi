var request = require('request');
var Excel = require('excel');
var exceldemolib = require('./exceldemolib.js')
var Excel = excel.Excel;
var OfficeExtension = require('office.runtime');

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

