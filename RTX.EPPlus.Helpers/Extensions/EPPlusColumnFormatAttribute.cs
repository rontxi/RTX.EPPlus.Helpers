using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTX.EPPlus.Helpers
{
  [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
  public class EPPlusColumnFormatAttribute : Attribute
  {
    public EPPlusColumnFormats format { get; set; }
    public EPPlusColumnTotalFormula total_formula { get; set; }
    public string date_format { get; set; }
    public string date_time_format { get; set; }
    public string currency_format { get; set; }
    public string percent_format { get; set; }
    public EPPPlusPivotTablePosition pivottable_position { get; set; }
    public DataFieldFunctions pivottable_function { get; set; }
    public ExcelHorizontalAlignment horizontal_alignment { get; set; }

    public EPPlusColumnFormatAttribute(
        EPPlusColumnFormats format = EPPlusColumnFormats.Default,
        EPPlusColumnTotalFormula total_formula = EPPlusColumnTotalFormula.None,
        string date_format = @"dd/mm/yyyy",
        string date_time_format = @"dd/mm/yyyy hh:mm",
        string currency_format = @"#,##0.00 €",
        string percent_format = @"#0.00%",
        EPPPlusPivotTablePosition pivottable_position =  EPPPlusPivotTablePosition.none,
        DataFieldFunctions pivottable_function = DataFieldFunctions.Count,
        ExcelHorizontalAlignment horizontal_alignment = ExcelHorizontalAlignment.General
    )
    {
      this.format = format;
      this.total_formula = total_formula;
      this.date_format = date_format;
      this.date_time_format = date_time_format;
      this.currency_format = currency_format;
      this.percent_format = percent_format;
      this.pivottable_position = pivottable_position;
      this.pivottable_function = pivottable_function;
      this.horizontal_alignment = horizontal_alignment;
    }
  }

  public enum EPPlusColumnFormats
  {
    Default = 0,
    Datetime = 1,
    Date = 2,
    Currency = 3,
    Percent = 4
  }

  public enum EPPlusColumnTotalFormula
  {
    None = 0,
    Average = 101,
    CountCellsWithNumbers = 102,
    CountAllCells = 103,
    Max = 104,
    Min = 105,
    Product = 106,
    Stdev = 107,
    Stdevp = 108,
    Sum = 109,
    Var = 110,
    Varp = 111
  }

  public enum EPPPlusPivotTablePosition
  {
    none = 0,
    dataField = 1,
    rowField = 2
  }
}
