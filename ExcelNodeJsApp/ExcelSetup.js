var request = require('request');
var excel = require('./excel.js');
var excelhelper = require('./excelhelper.js');

var Excel = excel.Excel;
var OfficeExtension = excel.OfficeExtension;

function initHttpRequestExecuteFunc(){
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
}

function getAccessToken(clientId, refreshToken){
	var tokenServiceUrl = "https://login.windows.net/common/oauth2/token";
    return OfficeExtension.HttpUtility.sendRequest(
        {
            url : tokenServiceUrl,
            method : "POST",
            body: "grant_type=refresh_token&refresh_token=" + refreshToken + "&client_id=" + clientId,
            headers: {
                "CONTENT-TYPE": "application/x-www-form-urlencoded"
            }
        })
        .then(function(resp){
            var v = JSON.parse(resp.body);
            return v["access_token"];
        });    
}

function initDefaultRequestUrlAndHeaders(newFile){
    
    var refreshToken = "AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJctMeWB1L1ZgevqQJAC4g5NCaOPNLn69ikpBeH1MjD7vE2ummRnokWTw3IN_YCFT1A4MGMdZN5d24hT4luOoLhb4Q93nw4XHOo6Xbvlr--u1wT9eJcK-fZ3xNmVnL6ywC8-j5icOXBul3ngw6fXbNbOLCEWrMnJYemDHchqzCbO0ldVmP0OgJdeFiQGg1VP3rJxrtX5tBOjH4nXRv8ZxPs_myBX6-sF2s9fsJyuPWv04NoJ2IJi9fzLYj7CLHAKfmFzWdlG_CorBu0NuIKiG7gLpo-2Md-UWXJzdV2OZ3bwdipbJgz0vKNmW4clHJ0P5h_7i3bax6Ql4E84klo2nyaqHAP_qOIvrBnRNmiBMHUJEG9USaXu_KXF3GrN3s5XpHEI2xGeqa6lu0M3G0-LuJBWblTNgOSna6LrhH-BFKAN7j28RELHNGRGCeIyHSQp5yVAj11ncHz63I8DoRixGnDCThJK0xXj4k4VY2WbBTrAfVBm4tLPS5lAbrAT7AzcOjes_Dt2EtSACUgxwtJjVPIkiJwPAOAMiT9VLJ5f6MiZNJgPIWmSVKqVJotZ-rhUvgORx_THvBETZcRCxgfXlDMjID8e69Ms2HlqpVKsMxA9_YnPpdL221-bphGfsw-L24sIwrbuRAQcn_JhAc-3zL_iAA";

	var clientId = "8563463e-ea18-4355-9297-41ff32200164";

    initHttpRequestExecuteFunc();

    var requestHeaders;

    return getAccessToken(clientId, refreshToken)
        .then(function(accessToken){
            requestHeaders = {Authorization: "Bearer " + accessToken};
            if (newFile){
                var date = new Date();
                var filename = "ShaoZhu" + date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + "-"
                    + date.getHours() + "-" + date.getMinutes() + "-" + date.getSeconds() + ".xlsx";
                return excelhelper.createBlankExcelFile(
                    "https://graph.microsoft.com/v1.0/me/drive/root",
                    filename,
                    requestHeaders);
            }
            else{
                return "https://graph.microsoft.com/v1.0/me/drive/root:/AgaveTest.xlsx:/workbook";
            }
        })
        .then(function(workbookUrl){
            return excelhelper.createSessionAndBuildUrlAndHeaders(workbookUrl, requestHeaders);
        })
        .then(function(requestUrlAndHeaders){
            OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders;
        });
}


exports.initDefaultRequestUrlAndHeaders = initDefaultRequestUrlAndHeaders;