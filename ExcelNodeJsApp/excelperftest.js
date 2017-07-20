var Excel = require('@microsoft/office-api/excel');
var exceldemolib = require('./exceldemolib.js')
var OfficeExtension = require('@microsoft/office-api/office.runtime');

// OfficeExtension.Utility._logEnabled = true;

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