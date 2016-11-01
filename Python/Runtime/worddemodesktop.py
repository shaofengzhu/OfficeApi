import runtime
import word
import worddemolib

if __name__ == "__main__":
    worddemolib.WordDemoLib.initDesktopContext()
    context = word.RequestContext()
    print("Insert image");
    worddemolib.WordDemoLib.insertSamplePictureAtEnd(context)
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = None
