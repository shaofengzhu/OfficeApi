var oauthhelper = require('./oauthhelper.js');
var excelhelper = require('./excelhelper.js');
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


var requestHeaders;

oauthhelper.getAccessToken(oauthhelper.clientId, oauthhelper.refreshToken)
    .then(function(accessToken){
        requestHeaders = {Authorization: "Bearer " + accessToken};
        var date = new Date();
        var filename = "ShaoZhu" + date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + "-"
            + date.getHours() + "-" + date.getMinutes() + "-" + date.getSeconds() + ".xlsx";
        return excelhelper.createBlankExcelFile(
            "https://graph.microsoft.com/v1.0/me/drive/root",
            filename,
            requestHeaders);
    })
    .then(function(workbookUrl){
        return excelhelper.createSessionAndBuildUrlAndHeaders(workbookUrl, requestHeaders);
    })
    .then(function(requestUrlAndHeaders){
        OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders;
    })
    .then(function(){
        return exceldemolib.dataPopulateSetup();
    })
    .then(function(){
        return exceldemolib.dataPopulateRun();
    })
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });



