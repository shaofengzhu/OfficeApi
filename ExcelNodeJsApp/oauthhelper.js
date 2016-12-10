var Excel = require('excel');
var OfficeExtension = require('office.runtime');

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

var refreshToken = "AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJPgPeMfFqm2lYBBJ7WxwrYMv54CMvD423ZlSn1riMrxtVIEnXFT1ckvRlQJcnlWAJXKOzCn1MaQOs-NOEGJl1qmJWCc7FypDCPwySixczXCyLrlyBFTWgUMSnVJC3SY5EOGlrWlrFzfTzqsZksgPy9EI1v1B-hg5tRHOkJ8X846T0altH1nOu0nZuGu-P6Pp9RpNORUDjipcePQg3fp5KMSJTM3OdfRYwBT9lUSdC7WzA_7xKdrS9eD6D24DUMx-uesXy24AVbNoq3juzLrFmHCcSQN1uPGuCJkU8c7yilWNYup0SNAfJtqDskQ_nnhCMRGOEnXnKzqC-c7CNC-0Kn0n-RTQ1N0ZV5UJ5yPhnxFJNdfwFXhkvoBC7qZv3QBot3vJBWeC1ayMymNK1q0B_-UMdmaNlrZJ7L1GFek3rUPDIK0E4l8ssoW_CwX5h6TPo9k51S2VU5z7LdeImjQjNs8ph5zEbH1xoGVo1vkupcBrBZWjA-jm9IwMj43tpXYTqnJYoNcR6e1DduH-56sxVQ9q8JRhrBmVRG6BR20CCgh6JYnpcf7ATgrfs4_NkIphYx9TkZOnlw5Nvr1FiArxRq46Q4nKU6JZn1QMl6imXaTalLfwkBklY8lCl466PXmS9UaPa2STLgJkU5RIX7cNdjCAA";

var clientId = "8563463e-ea18-4355-9297-41ff32200164";

exports.getAccessToken = getAccessToken;
exports.clientId = clientId;
exports.refreshToken = refreshToken;
