var request = require('request');
var excel = require('./excel.js');
var exceldemolib = require('./exceldemolib.js')
var Excel = excel.Excel;
var OfficeExtension = excel.OfficeExtension;

OfficeExtension.Utility._logEnabled = true;

OfficeExtension.HttpUtility.setCustomSendRequestFunc(function(req){
    return new OfficeExtension.Promise(function(resolve, reject){
        request(req, 
            function(err, resp){
                if (err){
                    reject(err);
                }
                else{
                    resolve(resp);
                }            
        });
    });
});


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

