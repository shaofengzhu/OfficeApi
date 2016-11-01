import runtime
import excel

def simpleRangeTest():
    context = excel.RequestContext("http://localhost:8052", runtime.RequestExecutionMode.immediateAndSlow)
    r = context.workbook.worksheets.getItem("Sheet1").getRange("A1:B2")
    print(r.values)

if __name__ == "__main__":
    simpleRangeTest()
