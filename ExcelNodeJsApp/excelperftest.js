var request = require('request');
var excel = require('./excel.js');
var exceldemolib = require('./exceldemolib.js')
var Excel = excel.Excel;
var OfficeExtension = excel.OfficeExtension;

// OfficeExtension.Utility._logEnabled = true;

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

var startTime = process.hrtime();
console.log(JSON.stringify(startTime));
exceldemolib.perfTest()
.then(function(){
    var diff = process.hrtime(startTime);
    console.log(JSON.stringify(diff));
    var milliSeconds = diff[0] * 1000 + diff[1] / 1000000;
    console.log("MilliSeconds:" + milliSeconds);
})
.catch(function(ex){
    console.log(JSON.stringify(ex));
});