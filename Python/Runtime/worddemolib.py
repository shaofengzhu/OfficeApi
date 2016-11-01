import numbers
import runtime
import word
import datetime
import json

class WordDemoLib:
    @staticmethod
    def initDesktopContext():
        requestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo();
        requestUrlAndHeaders.url = "http://localhost:8054";
        runtime.ClientRequestContext.defaultRequestUrlAndHeaders = requestUrlAndHeaders
        
    @staticmethod
    def insertPictureAtEnd(context: word.RequestContext, base64ImageData: str):
        context.document.body.insertInlinePictureFromBase64(base64ImageData, word.InsertLocation.end);
        context.sync();
        return
