using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBuilderDSL.Excel;

namespace testeA
{
    class Program
    {
        static void Main(string[] args)
        {
            var filename = "./Teste.xlsx";

            // for (uint i = 1; i < 100; i++)
            // {
            //     Console.Write(Util.GetCharColumn(i) + ",");    
            // }

            // Console.WriteLine(Util.GetCharColumn(26* 26 +5) + ",");//ZE
            // Console.WriteLine(Util.GetCharColumn(26* 26 + (26) +5) + ",");//AAE
            // Console.WriteLine(Util.GetCharColumn(26* 26 + (26 * 26) +5) + ",");//AZE
            // Console.WriteLine(Util.GetCharColumn(26* 26 + (26 * 26) + 26 + 1) + ",");//BAA


            ExcelBuilder.Builder()
                .WithPath(filename)
                .WithOverrideFile()
                .WithSheet(SheetBuilder.Builder()
                                .WithName("Turtles Sheet")
                                .WithHeader()
                                    .WithColumnName("Abc")
                                .EndHeader()
                                .Build()
                        )
                .WithSheet(SheetBuilder.Builder()
                                .WithName("Rafael")
                                .WithNewLine()
                                .WithColumnValue("teste")
                                .WithColumnValue(123)
                                .Build()
                        )
                .Build();
            // if(File.Exists(filename)){
            //     File.Delete(filename);
            // }
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
            //     Create(filename, SpreadsheetDocumentType.Workbook);

            // // Add a WorkbookPart to the document.
            // WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            // workbookpart.Workbook = new Workbook();

            // // Add a WorksheetPart to the WorkbookPart.
            // WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            // worksheetPart.Worksheet = new Worksheet(new SheetData());

            // // Add Sheets to the Workbook.
            // Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
            //     AppendChild<Sheets>(new Sheets());

            // // Append a new worksheet and associate it with the workbook.
            // Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.
            //     GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet abc" };
            // var sheetData = new SheetData();
            // var sheetData1 = new SheetData();


            // var row = new Row() { RowIndex = 1 };
            // var cell = new Cell() {CellReference = "A1",
            //      CellValue = new CellValue("Rafael Santos"),
            //      DataType  = new EnumValue<CellValues>(CellValues.String)
            // };
            // row.Append(cell);
            // sheetData.Append(row);

            // worksheetPart.Worksheet.Append(sheetData);

            // sheets.Append(sheet);


            // WorksheetPart worksheetPart1 = workbookpart.AddNewPart<WorksheetPart>();
            // worksheetPart1.Worksheet = new Worksheet(new SheetData());

            // var row1 = new Row() { RowIndex = 1 };
            // var cell1 = new Cell() {CellReference = "A1",
            //      CellValue = new CellValue("BlablaBla Santos"),
            //      DataType  = new EnumValue<CellValues>(CellValues.String)
            // };
            // row1.Append(cell1);
            // sheetData1.Append(row1);

            // worksheetPart1.Worksheet.Append(sheetData1);

            // sheets.Append(new Sheet() { Id = spreadsheetDocument.WorkbookPart.
            //     GetIdOfPart(worksheetPart1), SheetId = 2, Name = "Sheet 2" });

            // workbookpart.Workbook.Save();

            // // Close the document.
            // spreadsheetDocument.Close();

            Console.WriteLine("Hello World!");
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
    }
}
