using RTX.EPPlus.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTX.EPPlus.Tests
{
  public class ExportItem
  {
    [EPPlusColumnFormat(pivottable_position = EPPPlusPivotTablePosition.dataField)]
    public string column1 { get; set; }

    [DisplayName("INTT")]
    [EPPlusColumnFormat(total_formula = EPPlusColumnTotalFormula.Sum, pivottable_position = EPPPlusPivotTablePosition.rowField)]
    public int column2 { get; set; }

    [EPPlusColumnFormat(format = EPPlusColumnFormats.Datetime, total_formula = EPPlusColumnTotalFormula.CountAllCells)]
    public DateTime column3 { get; set; }

    [EPPlusColumnFormat(format= EPPlusColumnFormats.Currency, total_formula = EPPlusColumnTotalFormula.Sum)]
    public decimal column4 { get; set; }
    public decimal column5 { get; set; }

    public static List<ExportItem> GetTestData () {
        var lst = new List<ExportItem>();
      for (int i = 0; i < 100; i++)
      {
        lst.Add(new ExportItem() {
          column1 = ExportItem.RandomString(25),
          column2 = ExportItem.random.Next(1000, 9999),
          column3 = DateTime.Now.AddDays(ExportItem.random.Next(-100, +100)).AddHours(ExportItem.random.Next(-11, +11)).AddMinutes(ExportItem.random.Next(-59, +59)),
          column4 = (decimal)ExportItem.random.Next(10000, 99999) / 100,
          column5 = (decimal)ExportItem.random.Next(10000, 99999) / 100
        });
      }
      return lst;
    }

    private static Random random = new Random();

    private static string RandomString(int length)
    {
      const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
      return new string(Enumerable.Repeat(chars, length)
        .Select(s => s[random.Next(s.Length)]).ToArray());
    }
  }
}
