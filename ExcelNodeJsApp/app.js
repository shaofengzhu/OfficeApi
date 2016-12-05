var oauthhelper = require('./oauthhelper.js');
var excelhelper = require('./excelhelper.js');
var Excel = require('excel');
var exceldemolib = require('./exceldemolib.js');
var OfficeExtension = require('office.runtime');

OfficeExtension.Utility._logEnabled = true;

var session;

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
        session = new Excel.Session(workbookUrl, requestHeaders);
    })
    .then(function(){
        console.log("dataPopulateSetup");
        return exceldemolib.dataPopulateSetup(session);
    })
    .then(function(){
        console.log("dataPopulateRun");
        return exceldemolib.dataPopulateRun(session);
    })
    .then(function(){
        return session.close();
    })
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });



