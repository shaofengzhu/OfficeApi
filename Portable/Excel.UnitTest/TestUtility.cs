using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.ExcelServices;
using OfficeExtension;

namespace Excel.UnitTest
{
	static class TestUtility
	{
		public static void Dump2DArray(TestContext testContext, object[][] data)
		{
			for (int i = 0; i < data.Length; i++)
			{
				for (int j = 0; j < data[i].Length; j++)
				{
					testContext.WriteLine("({0}, {1})={2}", i, j, data[i][j]);
				}
			}
		}

		public static void DumpWorksheet(TestContext testContext, Worksheet sheet)
		{
			testContext.WriteLine("Sheet: Id={0}, Name={1}", sheet.Id, sheet.Name);
		}
	}
}
