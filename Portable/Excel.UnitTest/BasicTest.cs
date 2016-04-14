using System;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.ExcelServices;
using OfficeExtension;

namespace Excel.UnitTest
{
	[TestClass]
	public class BasicTest
	{
		public TestContext TestContext
		{
			get;
			set;
		}

		[TestMethod]
		public async Task GetWorksheets()
		{
			ExcelTestConfiguration testConfig = await ExcelTestConfiguration.GetInstance();
			ExcelRequestContext ctx = new ExcelRequestContext(testConfig.Url);
			testConfig.SetupRequestContext(ctx);
			ctx.Load(ctx.Workbook.Worksheets);
			await ctx.Sync();
			foreach(var worksheet in ctx.Workbook.Worksheets.Items)
			{
				TestUtility.DumpWorksheet(this.TestContext, worksheet);
			}
		}

		[TestMethod]
		public async Task GetRangeValue()
		{
			ExcelTestConfiguration testConfig = await ExcelTestConfiguration.GetInstance();
			ExcelRequestContext ctx = new ExcelRequestContext(testConfig.Url);
			testConfig.SetupRequestContext(ctx);
			Range r = ctx.Workbook.Worksheets.GetItem("Sheet1").GetRange("A1:B2");
			ctx.Load(r);
			await ctx.Sync();
			TestUtility.Dump2DArray(this.TestContext, r.Values);
		}

		[TestMethod]
		public async Task SetRangeValue()
		{
			ExcelTestConfiguration testConfig = await ExcelTestConfiguration.GetInstance();
			ExcelRequestContext ctx = new ExcelRequestContext(testConfig.Url);
			testConfig.SetupRequestContext(ctx);
			Range r = ctx.Workbook.Worksheets.GetItem("Sheet1").GetRange("A1:B2");
			r.Values = new object[][] { new object[] {"Hello", "World" }, new object[] { 100, 200 } };
			ctx.Load(r);
			await ctx.Sync();
			TestUtility.Dump2DArray(this.TestContext, r.Values);
		}

		[TestMethod]
		public async Task CreateDeleteSheet()
		{
			ExcelTestConfiguration testConfig = await ExcelTestConfiguration.GetInstance();
			ExcelRequestContext ctx = new ExcelRequestContext(testConfig.Url);
			testConfig.SetupRequestContext(ctx);
			string name = "S" + Guid.NewGuid().ToString("N").Substring(0, 8);
			Worksheet sheet = ctx.Workbook.Worksheets.Add(name);
			ctx.Load(sheet);
			await ctx.Sync();
			this.TestContext.WriteLine("Created sheet");
			TestUtility.DumpWorksheet(this.TestContext, sheet);
			string id = sheet.Id;
			sheet.Delete();
			ctx.Load(ctx.Workbook.Worksheets);
			await ctx.Sync();
			bool found = false;
			this.TestContext.WriteLine("After delete sheet {0}", id);
			foreach (Worksheet s in ctx.Workbook.Worksheets.Items)
			{
				TestUtility.DumpWorksheet(this.TestContext, s);
				if (s.Id == id)
				{
					found = true;
				}
			}

			Assert.IsFalse(found, "Expect not found={0}", found);
		}
	}
}
