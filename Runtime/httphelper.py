import sys
import json
import enum
import logging
import urllib.request


class RequestInfo:
    method = None
    url = None
    headers = {}
    body = None

class ResponseInfo:
    statusCode = None
    headers = {}
    body = None

class HttpUtility:
    @staticmethod
    def invoke(requestInfo: RequestInfo) -> ResponseInfo:
        if requestInfo.method is None:
            requestInfo.method = "GET"
        requestInfo.method = requestInfo.method.upper()
        if requestInfo.method == "GET" or requestInfo.method == "DELETE":
            requestInfo.body = None
        bodyData = None
        if requestInfo.body is not None:
            bodyData = requestInfo.body.encode("utf8")
        req = urllib.request.Request(requestInfo.url, method = requestInfo.method, headers = requestInfo.headers, data = bodyData);
        resp = urllib.request.urlopen(req)
        respInfo = ResponseInfo()
        respInfo.headers = {}
        for key, value in resp.info().items():
            respInfo.headers[key] = value
        respInfo.statusCode = resp.status
        charset = resp.info().get_content_charset()
        if charset is None:
            charset = "utf8"
        respInfo.body = resp.read().decode(charset)
        return respInfo

if __name__ == "__main__":
    req = RequestInfo()
    req.url = 'https://www.google.com/'
    response = HttpUtility.invoke(req)
    print(response.body)

