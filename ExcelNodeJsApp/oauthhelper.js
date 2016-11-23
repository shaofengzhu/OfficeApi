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

var refreshToken = "AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJl1v_-Gst20nb8P2_81OhMh1j32z9gDWkK6gxDdWOdf3gWsvSCFIl31sc-4LrVNFjBRtTP38qaFma5rbZaQ1YX6aXCAvvTfv-aKYTXcCkul4_WWdk8PeKl05oI_nRme2t4kUEy_NcDPecGsTyUfr9fp2B7YiGtwkYKXHO2Ypc9OXgE-ixti-gmpUEBFIDzylLjQg4oM9TbZnSu2h5iRnGY7Kt5lsRQie7XIdrNXTLIrCN9qg-gIjlHxKUquEETamg_Uau0mgkU7lXJvV7xymNvqx1cmXh7_jpazLnOHa_bbXxMN1LTQsx3sKZaiXlHiOehgyFqWMPEU-QaBD0Ki7ho8owmKgSZx0vUcx4t7tbszTHK7IkCQzjXFKhAr6LRERbP0Ycxz-GJR4ae66nKavoVzQtzIQbOWZdreG9tljou9vZY_te_IPgR4xlRtIfcWgRO3qtBuKrlL5ZuoDJFPbb6giaEWohQrwFMlG13pcCY7_11_x7yRndO_hT2mw2MVz5p3h4DG1W5ZkR_GClO_82GKeKCNtcf_h-vhKEv6FAr2JGKkyrRDAmIsvnuGQ2LnmgPSJa6MapQwrBYCuywqCsi0Pmzh89NqtOMUPtLFDNmvE6cb7b3aVCsNxE8-2E9lcCesy4u-Ol2jcX0q9n5tMWCiAA";

var clientId = "8563463e-ea18-4355-9297-41ff32200164";

exports.getAccessToken = getAccessToken;
exports.clientId = clientId;
exports.refreshToken = refreshToken;
