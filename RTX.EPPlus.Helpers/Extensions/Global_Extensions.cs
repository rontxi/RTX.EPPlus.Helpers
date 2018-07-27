using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using RTX.EPPlus.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace OfficeOpenXml
{
  public static class Global_Extensions
  {
    public static ExcelWorksheet AddWorkSheet(this ExcelPackage package, string WorksheetName)
    {
      return package.Workbook.Worksheets.Add(WorksheetName);
    }

    public static ExcelTable AddTable<T>(this ExcelWorksheet ws, string TableName, IEnumerable<T> data, TableStyles tableStyles = TableStyles.Light9, bool PrintHeaders = true, bool ShowTotal = false)
    {
      ws.Cells[1, 1].LoadFromCollection<T>(data, PrintHeaders);
      var table = ws.Tables.Add(ws.Dimension, TableName);
      table.ShowHeader = PrintHeaders;
      table.ShowTotal = ShowTotal;
      table.TableStyle = tableStyles;
      //Format columns
      var current_index = 1;
      foreach (var property in typeof(T).GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Instance))
      {
        var att = (EPPlusColumnFormatAttribute)Attribute.GetCustomAttribute(property, typeof(EPPlusColumnFormatAttribute));
        if (att != null)
        {
          setformat(ws, current_index, data.Count(), att, PrintHeaders);

          //Totals
          if (ShowTotal && att.total_formula != EPPlusColumnTotalFormula.None)
          {
            table.Columns[current_index - 1].TotalsRowFormula = string.Format("SUBTOTAL({0},[{1}])", (int)att.total_formula, table.Columns[current_index - 1].Name);
            if (att.format == EPPlusColumnFormats.Currency || att.format == EPPlusColumnFormats.Percent)
            {
              ws.Cells[ws.Dimension.End.Row, current_index].Style.Numberformat.Format = getFormat(att);
            }
          }
        }

        current_index++;
      }
      ws.Cells[ws.Dimension.Address].AutoFitColumns();
      return table;
    }

    public static ExcelPivotTable AddPivotTable<T>(this ExcelWorksheet ws, string TableName, ExcelRangeBase dataRange)
    {
      var pivotTable = ws.PivotTables.Add(ws.Cells[1, 1], dataRange, TableName);
      pivotTable.MultipleFieldFilters = true;
      pivotTable.RowGrandTotals = true;
      pivotTable.ColumGrandTotals = true;
      pivotTable.Compact = true;
      pivotTable.CompactData = true;
      pivotTable.GridDropZones = false;
      pivotTable.Outline = false;
      pivotTable.OutlineData = false;
      pivotTable.ShowError = true;
      pivotTable.ErrorCaption = "[error]";
      pivotTable.ShowHeaders = true;
      pivotTable.UseAutoFormatting = true;
      pivotTable.ApplyWidthHeightFormats = true;
      pivotTable.ShowDrill = true;
      pivotTable.FirstDataCol = 1;
      pivotTable.DataOnRows = false;
      pivotTable.RowHeaderCaption = TableName;

      foreach (var property in typeof(T).GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Instance))
      {
        var att = (EPPlusColumnFormatAttribute)Attribute.GetCustomAttribute(property, typeof(EPPlusColumnFormatAttribute));
        if (att != null)
        {
          if (att.pivottable_position != EPPPlusPivotTablePosition.none)
          {
            var attribute_display_name = (DisplayNameAttribute)Attribute.GetCustomAttribute(property, typeof(DisplayNameAttribute));
            var field_name = attribute_display_name != null ? attribute_display_name.DisplayName : property.Name;
            var field = pivotTable.Fields.FirstOrDefault(x => x.Name == field_name);
            if (field != null)
            {
              if (att.pivottable_position == EPPPlusPivotTablePosition.dataField)
              {
                var f = pivotTable.DataFields.Add(field);
                f.Function = att.pivottable_function;
                f.Name = field_name;
                f.Format = getFormat(att);
              }
              else if (att.pivottable_position == EPPPlusPivotTablePosition.rowField)
              {
                var f = pivotTable.RowFields.Add(field);
                f.Name = field_name;
              }
              //TODO PageFields && ColumnFields
            }
          }
        }
      }

      return pivotTable;
    }

    private static void setformat(ExcelWorksheet ws, int index, int data_rows_count, EPPlusColumnFormatAttribute att, bool PrintHeaders)
    {
      using (ExcelRange col = ws.Cells[PrintHeaders ? 2 : 1, index, data_rows_count + (PrintHeaders ? 1 : 0), index])
      {
        col.Style.Numberformat.Format = getFormat(att);
        col.Style.HorizontalAlignment = GetHorizontalAlignment(att);
      }
    }

    private static string getFormat(EPPlusColumnFormatAttribute att)
    {
      if (att.format == EPPlusColumnFormats.Date)
      {
        return att.date_format;
      }
      else if (att.format == EPPlusColumnFormats.Datetime)
      {
        return att.date_time_format;
      }
      else if (att.format == EPPlusColumnFormats.Currency)
      {
        return att.currency_format;
      }
      else if (att.format == EPPlusColumnFormats.Percent)
      {
        return att.percent_format;
      }
      else
      {
        return "";
      }
    }

    private static ExcelHorizontalAlignment GetHorizontalAlignment(EPPlusColumnFormatAttribute att)
    {
      if (att.format != EPPlusColumnFormats.Default && att.horizontal_alignment == ExcelHorizontalAlignment.General)
      {
        return ExcelHorizontalAlignment.Right;
      }
      else
        return att.horizontal_alignment;
    }
  }
}
