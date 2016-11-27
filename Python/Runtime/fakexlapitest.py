import runtime
import fakexlapi

class FakeExcelTest:
    @staticmethod
    def test_basic():
        ctx = fakexlapi.RequestContext()
        ctx.application.activeWorkbook.sheets.load()
        ctx.sync()
        print("Sheets:")
        for sheet in ctx.application.activeWorkbook.sheets.items:
            print(sheet.name)

if __name__ == "__main__":
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders = runtime.RequestUrlAndHeaderInfo()
    runtime.ClientRequestContext.defaultRequestUrlAndHeaders.url = "pipe://./fakeexcel/_api"

    methods = dir(FakeExcelTest)
    for method in methods:
        if method.startswith("test_"):
            print("invoke " + method)
            func = getattr(FakeExcelTest, method)
            func()

