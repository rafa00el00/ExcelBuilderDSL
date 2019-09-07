using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilderDSL.Excel
{

    public interface IExcelBuilder
    {
        IExcelBuilder WithPath(string path);
        IExcelBuilder WithOverrideFile();
        IExcelBuilder WithNonOverrideFile();
        IExcelBuilder WithSheet(Lazy<SheetClass> sheet);
        void Build();
    }
    public class ExcelBuilder : IExcelBuilder
    {
        private string Path { get; set; }
        public bool OverrideFile { get; private set; }
        private Queue<Lazy<SheetClass>> Sheets { get; set; }

        private ExcelBuilder()
        {
            OverrideFile = true;
            Sheets = new Queue<Lazy<SheetClass>>();
        }

        public static IExcelBuilder Builder()
        {
            return new ExcelBuilder();
        }


        public IExcelBuilder WithPath(string path)
        {
            this.Path = path;
            return this;
        }

        public IExcelBuilder WithOverrideFile()
        {
            this.OverrideFile = true;
            return this;
        }

        public IExcelBuilder WithNonOverrideFile()
        {
            this.OverrideFile = false;
            return this;
        }

        public IExcelBuilder Save()
        {
            return this;
        }

        public IExcelBuilder Close()
        {
            return this;
        }

        public IExcelBuilder WithSheet(Lazy<SheetClass> sheet)
        {
            Sheets.Enqueue(sheet);
            return this;
        }


        public void Build()
        {
            //Check If File Exist
            if (File.Exists(Path))
            {
                if (OverrideFile)
                    File.Delete(Path);
                else
                    throw new FileLoadException("File already exist, use WithOverride method for override exist file");
            }

            //Abre o documento para uso
            using (var spreadsheetDocument = SpreadsheetDocument.
                    Create(Path, SpreadsheetDocumentType.Workbook))
            {

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());



                ///
                /// Format

                var stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                stylesheet.AddNamespaceDeclaration("mc", "http: //schemas.openxmlformats.org/markup-compatibility/2006");
                stylesheet.AddNamespaceDeclaration("x14ac", "http: //schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                // create collections for fonts, fills, cellFormats, ...
                var fonts = new Fonts(); //{ Count = 1U, KnownFonts = true };
                var fills = new Fills();// { Count = 5U };
                var cellFormats = new CellFormats();// { Count = 4U };

                // create a font: bold, red, calibr
                Font font = new Font();
                font.Append(new FontSize() { Val = 21D });
                font.Append(new Color() { Rgb = "FF00FF" });
                //font.Append(new FontName() { Val = "Calibri" });
                //font.Append(new FontFamilyNumbering() { Val = 2 });
                //font.Append(new FontScheme() { Val = FontSchemeValues.Minor });
                font.Append(new Bold());
                // add the created font to the fonts collection
                // since this is the first added font it will gain the id 1U
                fonts.Append(font);

                // create a background: green
                Fill fill = new Fill();
                var patternFill = new PatternFill() { PatternType = PatternValues.Solid };
                patternFill.Append(new ForegroundColor() { Rgb = "00ff00" });
                patternFill.Append(new BackgroundColor() { Indexed = 64U });
                fill.Append(patternFill);
                fills.Append(fill);

                // create a cell format (combining font and background)
                // the first added font/fill/... has the id 0. The second 1,...
                cellFormats.AppendChild(new CellFormat() { FontId = 0, FillId = 0, ApplyFont = true, ApplyFill = true });

                // add the new collections to the stylesheet
                stylesheet.Append(fonts);
                stylesheet.Append(fills);
                stylesheet.Append(cellFormats);
                var stylePart = workbookpart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = stylesheet;
                stylePart.Stylesheet.Save();

                /// 
                /// 

                ///----------------///
                ///WorkSheet Logics///
                ///----------------///    

                uint i = 1;
                while (Sheets.Count > 0)
                {
                    var st = (Sheets.Dequeue() as Lazy<SheetClass>).Value;
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = st.Sheet;
                    sheets.Append(new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.
                     GetIdOfPart(worksheetPart),
                        SheetId = i,
                        Name = st.Name
                    });
                    i++;
                }



                workbookpart.Workbook.Save();



            }



        }

    }
}