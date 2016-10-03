var request = require('request');
var excel = require('./excel.js');

function sendHttpRequest(req){
    return new excel.OfficeExtension.Promise(function(resolve, reject){
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
}

function initHttpRequestExecuteFunc(){
    excel.OfficeExtension.setCustomHttpRequestExecuteFunc(sendHttpRequest);
}

function initDefaultRequestUrlAndHeaders(){
    var refreshToken = "AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJKdDNoyN1VN-o5tdTIZIjzkUJQQ1zKwBvXwS3HE1Grkl1YEDuy4C2BkJ1uXc7wmwZFKABl6qdbL4SS_bjckfreeJDA7S_Ild5xE_73zQPMo3tlPgKAaDKUhNXtMHtKj7QBGO-hyZzdczCJm3cQs195fOOWD_t_odi2ozFP7Op8y6Wb2IcNz3FtFvmsZXj82qyrOTbXrr0MTowS85cgZN3p1LVHuaOABKSoKdZh6hsy3kLkqYvkRcTf0ghPcD83FaWInj3mECfHCFq60KidmuWYOMAPJRK5MkQ0rykJmqH-8Ug2e4nFGmLe-CXklObNHlvtcQFXDKlEJ911T7zawbhNIcZ8Er2joyiDAC-nglCIRukKSc6XF5ud5vxSk___SKUPPIkHKgTOwxdoo84j3JpQoMPUswP3bxV_0KwG5mRCoDW3j3xp-QTHk8xTukfC3NCvIRlS0EIg6cJI8kh8RPZT7V9O26GLVPVCH2WdV2oZ8iR9DIJi9gdW4lJ3ZsmmFoVycg697vAVR49rMitRUEF51mngD2I04pI11ZFikmxYnZrHz-VyzVEojGJnJr7BMDHTETv3taThq5Co_pDXbIN5mknXU1X7sTI8lYjHvjbq5PQOPJxCE_Q38afI7py3yYfFhshjT5XUgEPimHyy2GKiSAA";
	var tokenServiceUrl = "https://login.windows.net/common/oauth2/token";
	var clientId = "8563463e-ea18-4355-9297-41ff32200164";

    var url = "https://graph.microsoft.com/v1.0/me/drive/root:/AgaveTest.xlsx:/workbook";
    var accessToken = "";
    var sessionId = "";

    return sendHttpRequest(
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
            accessToken = v["access_token"];
        })
        .then(function(){
			return sendHttpRequest(
				{
					url: url + "/createSession",
					method: "POST",
					body: JSON.stringify({ persistChanges: true }),
					headers: { Authorization: "Bearer " + accessToken }
				});            
        })
        .then(function(resp){
            var session = JSON.parse(resp.body);
            sessionId = session.id;

            excel.OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders = {
                url: url,
                headers: {
                    "Authorization": "Bearer " + accessToken,
                    "Workbook-Session-Id": sessionId
                }
            }            
        });
}

initHttpRequestExecuteFunc();

initDefaultRequestUrlAndHeaders()
    .then(function(){
        var ctx = new excel.Excel.RequestContext();
        var range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:B2');
        range.values = [["Oregon", "Washington"], [1234, "=A2 + 100"]];
        ctx.load(range);
        ctx.load(ctx.workbook.worksheets);
        ctx.sync()
        .then(function(){
            console.log(JSON.stringify(range.values));
            console.log(JSON.stringify(range.formulas));
            console.log("Worksheets");
            for (var i = 0; i < ctx.workbook.worksheets.items.length; i++){
                console.log(ctx.workbook.worksheets.items[i].name);
            }
        })
        .catch(function(ex){
            console.error(JSON.stringify(ex));
        });
    });
