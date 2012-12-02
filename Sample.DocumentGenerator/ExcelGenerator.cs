using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Sample.DocumentGenerator
{
  public class ExcelGenerator : IDocumentGenerator
  {
    public byte[] GenerateDocument(IDictionary<string, string> values, string fileName)
    {
      return GenerateDocument(values, fileName, 1000, "A", "B");
    }

    internal byte[] GenerateDocument(IDictionary<string, string> values, string fileName, uint rowStart, string columnKeysStart, string columnValuesStart)
    {
      if (values == null)
        throw new ArgumentException("Missing dictionary values");
      if (!File.Exists(fileName))
        throw new ArgumentException("File \"" + fileName + "\" do not exists");
      var tempFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
      return CreateFile(values, fileName, tempFileName, rowStart, columnKeysStart, columnValuesStart);
    }

    internal byte[] CreateFile(IDictionary<string, string> values, string fileName, string tempFileName, uint rowStart, string columnKeysStart, string columnValuesStart)
    {
      File.Copy(fileName, tempFileName);
      if (!File.Exists(tempFileName))
        throw new ArgumentException("Unable to create file: " + tempFileName);
      using (var spreadSheet = SpreadsheetDocument.Open(tempFileName, true))
      {
        var sheet = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
        if (sheet != null)
        {
          var worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheet.Id);
          SharedStringTablePart shareStringPart;
          if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
            shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
          else
            shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
          foreach (var value in values)
          {
            var index = InsertSharedStringItem(value.Key, shareStringPart);
            var cell = InsertCellInWorksheet(columnKeysStart, rowStart, worksheetPart);
            cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            index = InsertSharedStringItem(value.Value, shareStringPart);
            cell = InsertCellInWorksheet(columnValuesStart, rowStart, worksheetPart);
            cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            rowStart++;
          }
          worksheetPart.Worksheet.Save();
          //We force the calculation so the cells that link to this cell will update theirs values
          spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
          spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
        }
      }
      byte[] result = null;
      if (File.Exists(tempFileName))
      {
        result = File.ReadAllBytes(tempFileName);
        File.Delete(tempFileName);
      }
      return result;
    }

    //We use sharedstring becasue is more efficient. With same strings it use a ref and does not copy the value everytime
    internal static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
      if (shareStringPart.SharedStringTable == null)
        shareStringPart.SharedStringTable = new SharedStringTable();
      var i = 0;
      foreach (var item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
      {
        if (item.InnerText == text)
          return i;
        i++;
      }
      shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
      shareStringPart.SharedStringTable.Save();
      return i;
    }

    internal static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
      var worksheet = worksheetPart.Worksheet;
      var sheetData = worksheet.GetFirstChild<SheetData>();
      var cellReference = columnName + rowIndex;
      Row row;
      if (sheetData.Elements<Row>().Count(r => r.RowIndex == rowIndex) != 0)
        row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
      else
      {
        row = new Row() { RowIndex = rowIndex };
        sheetData.Append(row);
      }
      if (row.Elements<Cell>().Any(c => c.CellReference.Value == cellReference))
        return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
      var refCell = row.Elements<Cell>().FirstOrDefault(cell => String.Compare(cell.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase) > 0);
      var newCell = new Cell() { CellReference = cellReference };
      row.InsertBefore(newCell, refCell);
      worksheet.Save();
      return newCell;
    }
  }
}
