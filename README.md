
# EPPlus Library Helper

### Usage

Decorate whith attributes 

[DisplayName("Column Title")] for custom column name.

[EPPlusColumnFormat()] for:
* format 
* total_formula 
* date_format
* date_time_format
* currency_format
* pivottable_position
* pivottable_function
* horizontal_alignment

###Example code

```csharp
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
```
