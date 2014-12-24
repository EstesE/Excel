using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Excel2
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = "test.xlsx";
            using (var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                // DOM example
                //var workbookPart = spreadsheetDocument.WorkbookPart;
                //var worksheetPart = workbookPart.WorksheetParts.First();
                //var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                //string text;
                //foreach (Row r in sheetData.Elements<Row>())
                //{
                //    foreach (Cell c in r.Elements<Cell>())
                //    {
                //        text = c.CellValue.Text;
                //        Console.Write(text + " ");
                //    }
                //}
                //Console.ReadLine();


                // SAX example
                //var workbookPart = spreadsheetDocument.WorkbookPart;
                //var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();

                //var reader = OpenXmlReader.Create(worksheetPart);
                //string text;
                //while (reader.Read())
                //{
                //    if (reader.ElementType == typeof (CellValue))
                //    {
                //        text = reader.GetText();
                //        Console.Write(text + " ");
                //    }
                //}
                //Console.ReadLine();




                // Get specific cell's value
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>();
                foreach (Sheet sheet in sheets)
                {
                    var worksheetPart = (WorksheetPart) spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);
                    var worksheet = worksheetPart.Worksheet;

                    Cell cell = GetCell(worksheet, "D", 2);

                    Console.WriteLine(cell.CellValue.Text);

                    Console.ReadLine();
                }

            }


        }

        private static Cell GetCell(Worksheet worksheet, string columnName, int rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
            {
                return null;
            }

            return
                row.Elements<Cell>()
                    .Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0)
                    .First();
        }

        private static Row GetRow(Worksheet worksheet, int rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }
}
