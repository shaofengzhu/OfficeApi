var request = require('request');
var word = require('./word.js');
var worddemolib = require('./worddemolib.js')
var Word = word.Word;
var OfficeExtension = word.OfficeExtension;

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


OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = {url: "http://localhost:8054"};
worddemolib.insertSamplePictureAtEnd()
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });

