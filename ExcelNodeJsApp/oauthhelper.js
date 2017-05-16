var Excel = require('@microsoft/office-js/excel');
var OfficeExtension = require('@microsoft/office-js/office.runtime');

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

var refreshToken = "AQABAAAAAABnfiG-mA6NTae7CdWW7QfdueebNWYLYyQpgnp3T6jm4EL2lKbMeSCBEvo42QzN3vfyhaq_dzs-NBriv6inj3RAEPsS56e6JVvVaPJBuKPEluotgTuPcP9bFUYWkAMoyRlsE6HeXdHDcv2MmOAqZNkYUmghJCtckQX28oOL2urpdATpFjUpJutogx7uD6LPXZMwbCbbAX2CKeCTqgtoK1pm7tLRKJcovykfKGGFffFzBkI75NnGzVczrThuOAf72QJ5dNTqveDh0-cqCgGmqf4KEUjLMwfjX_4TYllZPxoZWtr2LJcOo_M9nj2RO1z29jsL5TUrBFGTtjRkWoTiCbgo9H2fHjY95wwQbPxCn7oyV-sb2B20FnDl1R4Q8pKDEnvTITHlJvTaL24FFLqQX61KI0o0c4uBV7sGRbUpSTsLrikMjEArFwDd5AugQsZ0USSJOfGaNqbiKqhm-P9ip-e20ROcmgjbqX1Fh4-shC-V3ZS0NduLMDeykxbE0JrY04bXHzV9pyWP85fjb6amFeYnUDY08VHrNmdHU5j3l1gpw9wkUAlKPXl3bMr7ZOD0QX867XzyvQV749rVynITf-lPMQNG26zHiSFFe4SwXiDY3NQIPihPip-OVRgTFi5a_N_8BPhLF595HBpneRqMgPSeyb-8s4QuQqbLgJrXN29gxwX35fjwaMaNWxzePK9uSqR2vLIBboqKKO0r9tvrQKzw2xFnh2apJsyZlpTwgq-QF0yDYwEH7maO9h-z2m3b9vIgAA";

var clientId = "8563463e-ea18-4355-9297-41ff32200164";

exports.getAccessToken = getAccessToken;
exports.clientId = clientId;
exports.refreshToken = refreshToken;
