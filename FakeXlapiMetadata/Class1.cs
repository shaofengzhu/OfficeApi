using System;
using System.Collections.Generic;
using Microsoft.OfficeExtension.CodeGen;
namespace FakeXlapi
{
    public interface Workbook
    {
        Worksheet ActiveWorksheet { get; }
    }

    public interface Worksheet
    {
        Range Range(string address);
    }

    public interface Range
    {
        string Text { get; set; }
    }
}
