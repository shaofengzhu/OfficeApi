var excel = require('./excel.js');

var Excel = excel.Excel;
var OfficeExtension = excel.OfficeExtension;

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

var refreshToken = "AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJFL7Jq7PHgd8yc-eECZcX_-xRWXl_oTMUXZk1KSqkVssXKKSbwFRlrsRwQVqoP4Ybkn8kW0iadaMr5dCZgqC9cpaTyUOxCrp9gMg4IjUZXRRJHLTfk2x-lNl0FBZCUztNgbPtlyjW8ak6Bc1uNTSaW4z9hChZJ6AH8X7hpjAgIfXBks2kf19oZk9no5BuwGQnl1p2k0mWfaYkMi-RiPhJec4g_Sn3i5gHx5DA4-yx-LbPfWMcOm-q_-kmH4cbUmibfpe4ST24tMoBmEb3LWbVJTCveAY4FB0fxqNTgQkudK8ocrC41QzMmha9u0J7-pnSWr6Nwwr2ry6MClhabLoBuxRru393PbTRMv5OJXBwvwrIFDu9TKeisk8ZvOa8jf5C3DQPNpl-UUl9wYSNv58qyaq6_aFYw8FsS3F0NoY_Tld6jYBcQsBGSpRQDNnourkOP56SNv0pCMQx6Z-zcYIGw-Iv30tB7WfNu9hE8bepsY4UqguNT4DvDrQmP8HaciDzVdLx8xgWDJyc8gWSd4UO2PWcUErmLUAkz3vYy907NrrvLlGH4OyXJ0CbCBCOksCET3_7bKdHuVGP_zDHP09hKHAAIolF95fj0D0iF1hpHIoqtaSNL2RCymFs89d5ZGqaYKbwfaCRS17ylGyD04YSXCAA";

var clientId = "8563463e-ea18-4355-9297-41ff32200164";

exports.getAccessToken = getAccessToken;
exports.clientId = clientId;
exports.refreshToken = refreshToken;
