var request = require('request');
var word = require('./word.js');
var worddemolib = require('./worddemolib.js')
var Word = word.Word;
var OfficeExtension = word.OfficeExtension;

OfficeExtension.Utility._logEnabled = true;

OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = {url: "http://localhost:8054"};
worddemolib.insertSamplePictureAtEnd()
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });

