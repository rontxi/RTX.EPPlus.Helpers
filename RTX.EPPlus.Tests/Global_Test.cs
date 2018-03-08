using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace RTX.EPPlus.Tests
{
  [TestClass]
  public class Global_Test
  {
    [TestMethod]
    public void FirstTest()
    {
      using (ExcelPackage pck = new ExcelPackage())
      {
        var ws_dades = pck.AddWorkSheet("DataSheet");
        ws_dades.AddTable<ExportItem>("Table1", ExportItem.GetTestData(), ShowTotal:true);
        var ws_promotor = pck.AddWorkSheet("PivottableSheet");
        ws_promotor.AddPivotTable<ExportItem>("PivotTable", ws_dades.Cells[ws_dades.Dimension.Address]);

        string path = @"C:\temp\test1.xlsx";
        Stream stream = File.Create(path);
        pck.SaveAs(stream);
        stream.Close();
      }
    }
  }
}
