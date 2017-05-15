var oauthhelper = require('./oauthhelper.js');
var excelhelper = require('./excelhelper.js');
var request = require('request');
var Excel = require('excel');
var exceldemolib = require('./exceldemolib.js');
var OfficeExtension = require('office.runtime');
var fetch = require('node-fetch');

OfficeExtension.Utility._logEnabled = true;

var session;

OfficeExtension.HttpUtility.setCustomSendRequestFunc(function(req){
    return fetch(req.url, {method: req.method, headers: req.headers, body: req.body})
    .then(function(resp){
        return {statusCode: resp.status, headers: resp.headers, body: resp.body};
    });
});

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
        return exceldemolib.dataPopulateSetup(session);
    })
    .then(function(){
        return exceldemolib.dataPopulateRun(session);
    })
    .then(function(){
        return session.close();
    })
    .catch(function(ex){
        console.error(JSON.stringify(ex));
    });



