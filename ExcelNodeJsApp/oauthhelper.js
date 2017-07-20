var Excel = require('@microsoft/office-api/excel');
var OfficeExtension = require('@microsoft/office-api/office.runtime');

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

var refreshToken = "AQABAAAAAABnfiG-mA6NTae7CdWW7QfdKhRuWfWMz5ngqvkdGmdPLrwkorDjtlkW1vCkGiK7gyJ_kG5b_dFutRUpYHnA97VWXfyQPpB89uR2Bjjdc8-qnRVJOk9O9VW8ReaYvI7LMusbW1OBd0L4Qbas4EnG5lnrkiWCacR8GeamzIuK1bdJg797BvGG5OeOiNEgr9KNIoVM15DhJgkt4FVkxmyQC1g2HdsNaMNy6zn6GCxNB5hR5aPY9leJ-cjNWfhLP8---FnQqE3LM2JLVoIsShjqMQJvKhLWQFK_CSDcs4xXJ7QVGUW96V1cpbv9HuoQZGchaQvLpe7ygPTQiaBN5fljOZP3cQWLxYJpZb7NKbNKs19RkuTUOo1sZQ8raPX5JQcbRWbMl3Fp4QBs_RfMHsvoQ91eQMyV5JehQinliLtPOIn_HKSJdbb7Y-70Zv6fzjWd6QpFPNmCMn_FGTH496uPk4HeZv-32uYooA4yLyb9FFAJm9Aas95nsXeVgREKgi2jyRwRrne4YOzE7utCY2V8hZb4szsGKSvRQxq7W2V-O0CAcVpijEew1F_4jkeQbEQcB_2Pi9v3LZOl3YKM1q1Fu3PSMwsxDYfTl7koqHtvFhhfbM86KQP5D3uluc5XankWNe9xdmZ2_jZ-xf3pou1A0xrVbTs-ycUEWdXOYvD-1iPzPvNgrXa7lRnBwjVaVWbL6yBzs859Hf4fUx-K7Cj6D8iJx02vbsDOfWnhA8VP9NcBrji5zxb1-w1DdqBbmg83bZYgAA";

var clientId = "8563463e-ea18-4355-9297-41ff32200164";

exports.getAccessToken = getAccessToken;
exports.clientId = clientId;
exports.refreshToken = refreshToken;
