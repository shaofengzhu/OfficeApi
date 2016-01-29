import sys
import json
import random
import excel

class ExcelTest:
    serverUrl = "http://shaozhu-ttvm8.redmond.corp.microsoft.com/th/WacRest.ashx/transport_wopi/Application_Excel/wachost_/Fi_anonymous~AgaveTest.xlsx/ak_1%7CGN=R3Vlc3Q=&SN=OTYwMjY3MTM0&IT=NTI0NzU4MjM3MjMzNzY2MTg3Mg==&PU=OTYwMjY3MTM0&SR=YW5vbnltb3Vz&TZ=MTExOQ==&SA=RmFsc2U=&LE=RmFsc2U=&AG=VHJ1ZQ==&RH=nNbTL6fvW2u38x1-jhY2YJ2RiYya97tuj6UnTFEfsD8=/_api"

    @staticmethod
    def setupRequestContext(ctx: excel.RequestContext):
        pass

    @staticmethod
    def clearWorksheet(ctx: excel.RequestContext, sheetName: str):
        sheet = ctx.workbook.worksheets.getItem(sheetName)
        ctx.load(sheet.tables)
        ctx.sync()
        for table in sheet.tables.items:
            table.delete()
        sheet.getRange(None).clear(excel.ClearApplyTo.all)
        ctx.sync()

    @staticmethod
    def test_Range_SetValueReadValue():
        ctx = excel.RequestContext(ExcelTest.serverUrl)
        ExcelTest.setupRequestContext(ctx)
        r = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2")
        r.values = [["Hello", "World"], [100, 200]]
        r.load()
        ctx.sync()
        print(r.values)
        print(r.address)
        print(r)
        
    @staticmethod
    def test_Worksheet_GetWorksheetCollection():
        ctx = excel.RequestContext(ExcelTest.serverUrl)
        ExcelTest.setupRequestContext(ctx)
        ctx.workbook.worksheets.load()
        ctx.sync()
        print("Worksheets")
        for sheet in ctx.workbook.worksheets.items:
            print(sheet.name)

    @staticmethod
    def test_Worksheet_AddDeleteWorksheet():
        ctx = excel.RequestContext(ExcelTest.serverUrl)
        ExcelTest.setupRequestContext(ctx)
        random.seed()
        name = "PythonTest" + str(random.randint(1, 3000))
        sheet = ctx.workbook.worksheets.add(name)
        sheet.load()
        ctx.sync()
        print("Created sheet " + sheet.name)
        print("Created sheetIndex " + str(sheet.position))
        sheet.delete()
        ctx.sync()
        print("Deleted sheet")

    @staticmethod
    def test_Table_GetCollection():
        ctx = excel.RequestContext(ExcelTest.serverUrl)
        ExcelTest.setupRequestContext(ctx)
        ctx.load(ctx.workbook.tables)
        ctx.sync()
        for table in ctx.workbook.tables.items:
            print(table.name)

    @staticmethod
    def test_Table_CreateTable():
        sheetName = "Tables"
        tableAddress = sheetName + "!A23:B25"
        ctx = excel.RequestContext(ExcelTest.serverUrl)
        ExcelTest.setupRequestContext(ctx)
        ExcelTest.clearWorksheet(ctx, sheetName)
        t = ctx.workbook.tables.add(tableAddress, True)
        ctx.load(t)
        ctx.sync()
        print("Created table id=" + str(t.id))
        print("Created table name=" + t.name)

    @staticmethod
    def test_Table_CreateDeleteTable():
        sheetName = "Tables"
        tableAddress = sheetName + "!A23:B25"
        ctx = excel.RequestContext(ExcelTest.serverUrl)
        ExcelTest.setupRequestContext(ctx)
        ExcelTest.clearWorksheet(ctx, sheetName)
        t = ctx.workbook.tables.add(tableAddress, True)
        ctx.load(t)
        ctx.sync()
        print("Created table id=" + str(t.id))
        print("Created table name=" + t.name)
        t.delete()
        ctx.sync()
        print("Deleted table")

if __name__ == "__main__":
    methods = dir(ExcelTest)
    for method in methods:
        if method.startswith("test_"):
            print("invoke " + method)
            func = getattr(ExcelTest, method)
            func()

