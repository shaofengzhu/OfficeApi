import sys
import json
import httphelper

class GraphFileAccessInfo:
    def __init__(self):
        self.fileId = ""
        self.accessToken = ""
        self.fileWorkbookUrl = ""


class OAuthUtility:

    @staticmethod
    def getFileAccessInfo(useProductionEnvironment: bool, filename: str) -> GraphFileAccessInfo:
        graphRootUrl = ""
        if  useProductionEnvironment:
            graphRootUrl = "https://graph.microsoft.com/testexcel"
        else:
            graphRootUrl = "https://graph.microsoft-ppe.com/testexcel"
        accessToken = OAuthUtility.getAccessToken(useProductionEnvironment)
        requestInfo = httphelper.RequestInfo()
        requestInfo.method = "GET"
        requestInfo.url = graphRootUrl + "/me/drive/root/children"
        requestInfo.headers["Authorization"] = "Bearer " + accessToken
        responseInfo = httphelper.HttpUtility.invoke(requestInfo)
        if responseInfo.statusCode != 200:
            raise RuntimeError("Cannot get files")
        resp = json.loads(responseInfo.body)
        files = resp.get("value")
        fileId = ""
        for file in files:
            if file.get("name") is not None and file.get("name").upper() == filename.upper():
                fileId = file.get("id")
                break
        if len(fileId) == 0:
            raise RuntimeError("Cannot find file")

        ret = GraphFileAccessInfo()
        ret.fileId = fileId
        ret.accessToken = accessToken
        ret.fileWorkbookUrl = graphRootUrl + "/me/drive/items/" + fileId + "/workbook"
        return ret

    @staticmethod
    def getAccessToken(useProductionEnvironment: bool) -> str:
        tokenServiceUrl = ""
        clientId = ""
        refreshToken = ""
        if useProductionEnvironment:
            tokenServiceUrl = "https://login.windows.net/common/oauth2/token"
            clientId = "8563463e-ea18-4355-9297-41ff32200164"
            refreshToken = "AQABAAAAAADRNYRQ3dhRSrm-4K-adpCJ8KFNC8qd13f7CP1WIcVSUQEiBZD-Z88D0muKYiHFiWn-ghYddJstawWhhfUWs6OdRSkPTrEoOjku6-qlp8H4g0eyC7v-_EaJVWOL4lCgUoHU_mZm_xUgdj5RX3cmoSHJgPmwj4sWxZr0L-ae6eUoprue7YqI8YhM4fqSx26iCCPXtaSGdagRGMkrtvU3dfe72-ecSWdo6Neq93HJirmdvwu0ILFWFw42f26XX0QQCssx_B2LRxi8NNaznUaPS1h6OITm_WDQeNzGnntrfya4cJNxZXxk5ZfWh5GBDGXLk1kfQWyLOgpFHddv9jiKL5tXdR9m7WZyqKyIjGnuc9PFKGRiTd69VZ2eSaaNwkqMUlpILdX683yGWy9yMuESHNLdt0sk_nXb1LMYawvlYgezCoHS3UYAxn42XyVM4VBpKNeiL6-dLqETyn1KXYyWgh40fxNQfYfSnUx5yyUYxGdh2n8H5NylW61j9RO2aYPY_Bl6kUJbqF62w83PlSZahZf3SiQJtI9yCvhHGCbW5tCgTHGYk3nr7wTIILso-PGosGZAyi_cGjYhhTeb0HZFR-aZilThNplGDKAPmolhM4pIcfA0EcPWQ5Ar0uYjot1aLxQ6zpmhjAv5CytrT78KRHijQObuvjVWMwbvN5tvYjJf0SAA"
        else:
            tokenServiceUrl = "https://login.windows-ppe.net/common/oauth2/token"
            clientId = "09d9cc54-6048-4c79-b468-99aa29c6e98d"
            refreshToken = "AAABAAAAo3ZCPl0FaU2WWRdLWLHperA8sJ4PqXDxCTLjPNRJsutVXPEEEc-q4h3YgZ2IUx9ogcH0iUE7juPkQGt_9kW7UIKmhfoye0ob3Y629xtAFc20jv3mO1cSQlKzuaPjjwIg91RQ1MbKbBqVLKeWRJ62MYJoBH4pnsLQXbv_H4hpENnIfT4CKSbDA4MCKhjXzL1TyCBSAFfjU-5ddUvyj_m2HkIL0mdysjkDpLY4cMNr1gBVxW4isHYkR23pGZsVJdVgJgCJ_k4Gf49Pypzlor6qSynu3w9TtlEZsKswMLFqKKNqnMYJh6eSLh7Q3ljXW21iDmsxXaT-BTiuBwrJN4if3oRHyVbo4IeNHzc3dHrsBjlfkR8LdhrdPvoZz9OD7RYaopaN-mAtZplN16I-pev_ii6Y73FCPp3yKDXNoIhJC2O-Wcgl8Ev0CPOeSq8tdtfE-VE53SIgZnc0MjE4WiZzFyejzatXDIhI9XQAXJC5JPGhL1q6AYtoP4Zih_sLDywxitrU9XikneZyjy1RGmmxMzuOjyafXZnlTLLD7ko7XYADZNps7J4GW2FSeCOiOEvAIAA"
        return OAuthUtility._getAccessTokenFromRefreshToken(clientId, tokenServiceUrl, refreshToken)


    @staticmethod
    def _getAccessTokenFromRefreshToken(clientId : str, tokenSvcUrl : str, refreshToken : str) -> str:
        requestInfo = httphelper.RequestInfo()
        requestInfo.url = tokenSvcUrl
        requestInfo.method = "POST"
        requestInfo.body = "grant_type=refresh_token&refresh_token=" + refreshToken + "&client_id=" + clientId
        requestInfo.headers = {}
        requestInfo.headers["CONTENT-TYPE"] = "application/x-www-form-urlencoded"
        responseInfo = httphelper.HttpUtility.invoke(requestInfo)
        if responseInfo.statusCode != 200:
            raise RuntimeError("Unable to get token")
        resp = json.loads(responseInfo.body)
        accessToken = resp.get("access_token")
        return accessToken

