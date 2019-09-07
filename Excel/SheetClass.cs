using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilderDSL.Excel
{
    public class SheetClass
    {
        public string Name { get; set; }
        public Worksheet Sheet { get; set; }
    }
}
