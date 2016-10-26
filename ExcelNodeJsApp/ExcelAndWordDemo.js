var request = require('request');
var excel = require('./excel.js');
var word = require('./word.js');
var exceldemolib = require('./exceldemolib.js');
var worddemolib = require('./worddemolib.js');
var childprocess = require('child_process');
var fs = require('fs');
var Excel = excel.Excel;
var Word = word.Word;
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

var bitmap = fs.readFileSync('blank.xlsx');
var buf = new Buffer(bitmap);
fs.writeFileSync('demo.xlsx', buf);

bitmap = fs.readFileSync('blank.docx');
buf = new Buffer(bitmap);
fs.writeFileSync('demo.docx', buf);

var chartImageBase64 = "";

var p = new OfficeExtension.Promise(function(resolve, reject){
    resolve(null);
});
    p
    // .then(function(){
    //     // launch Excel and Word
    //     var p1 = childprocess.exec("demo.xlsx", function(err, stdout, stderr){
    //         if (err){
    //             console.log(JSON.stringify(err));
    //         }
    //     });
    //     var p2 = childprocess.exec("demo.docx", function(err, stdout, stderr){
    //         if (err){
    //             console.log(JSON.stringify(err));
    //         }
    //     });

    //     return new OfficeExtension.Promise(function(resolve, reject){
    //         setTimeout(function(){
    //             resolve(null);
    //         }, 60 * 1000);
    //     });
    // })
    .then(function(){
        // set context to Excel
        console.log("Set context to Excel");
        OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = {url: "http://localhost:8052"};
    })
    .then(function(){
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
        chartImageBase64 = imageData;
    })
    .then(function(){
        // set context to Word
        console.log("Set context to Word");
        OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = {url: "http://localhost:8054"};
    })
    .then(function(){
        return worddemolib.insertPictureAtEnd(chartImageBase64);
    })
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });

